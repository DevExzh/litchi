//! Checked row and column sizing values for Numbers tables.

use std::fmt;

/// Exact-source row and column size transactions.
pub mod transaction;

/// One physical table axis addressed by a zero-based index.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub enum Dimension {
    /// A row at the given zero-based index.
    Row(usize),
    /// A column at the given zero-based index.
    Column(usize),
}

impl Dimension {
    /// Returns the zero-based index within the selected axis.
    #[must_use]
    pub const fn index(self) -> usize {
        match self {
            Self::Row(index) | Self::Column(index) => index,
        }
    }

    /// Returns the singular axis name used in diagnostics.
    #[must_use]
    pub const fn noun(self) -> &'static str {
        match self {
            Self::Row(_) => "row",
            Self::Column(_) => "column",
        }
    }
}

/// A validated positive, finite table dimension measured in points.
#[derive(Clone, Copy, Debug, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct Points(f32);

impl Points {
    /// Validates and constructs a point measurement.
    ///
    /// Values must be finite and strictly positive. Native zero is reserved
    /// for [`Size::Default`] and is not representable as explicit sizing.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for zero, negative, infinite, or NaN input.
    pub fn new(value: f32) -> Result<Self, Error> {
        if !value.is_finite() || value <= 0.0 {
            return Err(Error::Invalid);
        }
        Ok(Self(value))
    }

    /// Returns the point measurement.
    #[must_use]
    pub const fn value(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for Points {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self, Self::Error> {
        Self::new(value)
    }
}

impl From<Points> for f32 {
    fn from(points: Points) -> Self {
        points.value()
    }
}

/// Either the table style's default size or an explicit point override.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum Size {
    /// Use the table style's native default.
    #[default]
    Default,
    /// Use an explicit positive point measurement.
    Points(Points),
}

impl Size {
    /// Validates and constructs an explicit point override.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Invalid`] for zero, negative, infinite, or NaN input.
    pub fn points(value: f32) -> Result<Self, Error> {
        Ok(Self::Points(Points::new(value)?))
    }
}

/// Errors returned while constructing a dimension value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Error {
    /// A point measurement was zero, negative, infinite, or NaN.
    Invalid,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid => {
                formatter.write_str("table dimension points must be positive and finite")
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
    fn dimension_is_compact_and_lossless() {
        assert_eq!(Dimension::Row(3).index(), 3);
        assert_eq!(Dimension::Column(7).index(), 7);
        assert_eq!(Dimension::Row(3).noun(), "row");
        assert_eq!(Dimension::Column(7).noun(), "column");
    }

    #[test]
    fn points_reject_non_positive_and_non_finite_values() {
        for value in [0.0, -1.0, f32::INFINITY, f32::NEG_INFINITY, f32::NAN] {
            assert!(Points::new(value).is_err());
            assert!(Size::points(value).is_err());
        }
        assert_eq!(Points::new(32.0).unwrap().value(), 32.0);
    }

    #[test]
    fn size_distinguishes_default_from_explicit_points() {
        assert_eq!(Size::default(), Size::Default);
        assert_eq!(
            Size::points(32.0).unwrap(),
            Size::Points(Points::new(32.0).unwrap())
        );
    }

    #[test]
    fn points_has_no_allocation_or_hidden_state() {
        assert_eq!(size_of::<Points>(), size_of::<f32>());
        assert!(size_of::<Dimension>() <= 2 * size_of::<usize>());
    }
}
