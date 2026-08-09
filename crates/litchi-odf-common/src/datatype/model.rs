//! Semantic `ODF` scalar models.

use std::fmt;

// ============================================================================
// BOOLEAN, DATE, AND DATETIME MARKERS
// ============================================================================

/// Boolean data type conversion utilities.
///
/// Converts between `ODF` boolean format (`"true"`/`"false"`) and Rust `bool`.
pub struct Boolean;

/// Date data type conversion utilities.
///
/// Converts between `ODF` date format (`YYYY-MM-DD`) and `chrono::NaiveDate`.
pub struct Date;

/// `DateTime` data type conversion utilities.
///
/// Converts between `ODF` datetime format and `chrono::DateTime` values.
pub struct DateTime;

/// Duration data type conversion utilities.
///
/// Converts between `ODF` duration format and `chrono::Duration`, with
/// [`DurationValue`] available when the complete XML Schema duration must be
/// retained.
pub struct Duration;

/// Exact XML Schema duration value used by `ODF`.
///
/// Calendar years and months cannot be represented by [`chrono::Duration`]
/// without a reference date. This type retains every component and its exact
/// lexical representation, including arbitrary-width integers and fractional
/// seconds.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct DurationValue {
    pub(super) lexical: String,
    pub(super) negative: bool,
    pub(super) years: Option<String>,
    pub(super) months: Option<String>,
    pub(super) days: Option<String>,
    pub(super) hours: Option<String>,
    pub(super) minutes: Option<String>,
    pub(super) seconds: Option<String>,
}

impl DurationValue {
    /// Return the exact validated `ODF` lexical representation.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.lexical
    }

    /// Whether the duration carries a negative sign.
    #[must_use]
    pub fn is_negative(&self) -> bool {
        self.negative
    }

    /// Calendar-year component, if present.
    #[must_use]
    pub fn years(&self) -> Option<&str> {
        self.years.as_deref()
    }

    /// Calendar-month component, if present.
    #[must_use]
    pub fn months(&self) -> Option<&str> {
        self.months.as_deref()
    }

    /// Day component, if present.
    #[must_use]
    pub fn days(&self) -> Option<&str> {
        self.days.as_deref()
    }

    /// Hour component, if present.
    #[must_use]
    pub fn hours(&self) -> Option<&str> {
        self.hours.as_deref()
    }

    /// Minute component, if present.
    #[must_use]
    pub fn minutes(&self) -> Option<&str> {
        self.minutes.as_deref()
    }

    /// Seconds component, including its fractional part, if present.
    #[must_use]
    pub fn seconds(&self) -> Option<&str> {
        self.seconds.as_deref()
    }
}

impl fmt::Display for DurationValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.lexical)
    }
}
