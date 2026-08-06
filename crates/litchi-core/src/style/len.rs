use std::fmt;

use crate::unit::{EMUS_PER_CM, EMUS_PER_INCH};

/// Length measurement with units.
///
/// Represents a measurement value used for dimensions, positions, etc.
/// Office formats primarily use EMUs (English Metric Units).
///
/// # Examples
///
/// ```rust
/// use litchi_core::Length;
///
/// // Create from EMUs
/// let length = Length::from_emus(914400); // 1 inch
///
/// // Convert to different units
/// let inches = length.inches();
/// let cm = length.cm();
/// ```
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Length {
    /// Value in EMUs (English Metric Units)
    /// 1 inch = 914,400 EMUs
    /// 1 cm = 360,000 EMUs
    emus: i64,
}

impl Length {
    /// Create a length from EMUs (English Metric Units).
    ///
    /// EMUs are the native unit used in Office Open XML formats.
    /// - 1 inch = 914,400 EMUs
    /// - 1 cm = 360,000 EMUs
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_core::Length;
    ///
    /// let length = Length::from_emus(914400); // 1 inch
    /// ```
    #[inline]
    #[must_use]
    pub const fn from_emus(emus: i64) -> Self {
        Self { emus }
    }

    /// Create a length from inches.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_core::Length;
    ///
    /// let length = Length::from_inches(1.0);
    /// ```
    #[inline]
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "float-to-int truncation toward zero is the intended conversion semantics"
    )]
    #[allow(
        clippy::cast_precision_loss,
        reason = "EMUS_PER_INCH (914_400) is exactly representable as f64"
    )]
    pub fn from_inches(inches: f64) -> Self {
        Self {
            emus: (inches * EMUS_PER_INCH as f64) as i64,
        }
    }

    /// Create a length from centimeters.
    ///
    /// # Examples
    ///
    /// ```rust
    /// use litchi_core::Length;
    ///
    /// let length = Length::from_cm(2.54); // ~1 inch
    /// ```
    #[inline]
    #[must_use]
    #[allow(
        clippy::cast_possible_truncation,
        reason = "float-to-int truncation toward zero is the intended conversion semantics"
    )]
    #[allow(
        clippy::cast_precision_loss,
        reason = "EMUS_PER_CM (360_000) is exactly representable as f64"
    )]
    pub fn from_cm(cm: f64) -> Self {
        Self {
            emus: (cm * EMUS_PER_CM as f64) as i64,
        }
    }

    /// Get the value in EMUs.
    #[inline]
    #[must_use]
    pub const fn emus(&self) -> i64 {
        self.emus
    }

    /// Convert to inches.
    #[inline]
    #[must_use]
    #[allow(
        clippy::cast_precision_loss,
        reason = "f64 precision is sufficient for document dimensions; EMUS_PER_INCH is exactly representable"
    )]
    pub fn inches(&self) -> f64 {
        self.emus as f64 / EMUS_PER_INCH as f64
    }

    /// Convert to centimeters.
    #[inline]
    #[must_use]
    #[allow(
        clippy::cast_precision_loss,
        reason = "f64 precision is sufficient for document dimensions; EMUS_PER_CM is exactly representable"
    )]
    pub fn cm(&self) -> f64 {
        self.emus as f64 / EMUS_PER_CM as f64
    }

    /// Convert to points (1/72 inch).
    #[inline]
    #[must_use]
    pub fn points(&self) -> f64 {
        self.inches() * 72.0
    }
}

impl fmt::Display for Length {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "{:.2}\"", self.inches())
    }
}
