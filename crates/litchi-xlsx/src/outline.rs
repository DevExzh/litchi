//! Checked worksheet outline levels shared by rows and columns.

use thiserror::Error;

/// Checked worksheet outline level in Office's `0..=7` domain.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Outline(u8);

impl Outline {
    /// No outline grouping.
    pub const NONE: Self = Self(0);

    /// Validate one outline level.
    pub const fn new(value: u8) -> Result<Self, OutlineError> {
        if value <= 7 {
            Ok(Self(value))
        } else {
            Err(OutlineError {
                value: value as i64,
            })
        }
    }

    /// Return the checked numeric level.
    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// Invalid Office worksheet outline level.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[error("outline level {value} is outside the Office range 0..=7")]
pub struct OutlineError {
    value: i64,
}

impl OutlineError {
    /// Rejected numeric value.
    #[must_use]
    pub const fn value(self) -> i64 {
        self.value
    }
}

/// Convenient checked-or-raw input for [`Outline`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum OutlineAt {
    /// A level that has already been checked.
    Checked(Outline),
    /// A raw level validated when resolved.
    Level(i64),
}

impl OutlineAt {
    /// Resolve this input into a checked outline level.
    pub const fn resolve(self) -> Result<Outline, OutlineError> {
        match self {
            Self::Checked(level) => Ok(level),
            Self::Level(value) if value < 0 || value > 7 => Err(OutlineError { value }),
            Self::Level(value) => Outline::new(value.to_le_bytes()[0]),
        }
    }
}

impl From<Outline> for OutlineAt {
    fn from(value: Outline) -> Self {
        Self::Checked(value)
    }
}

impl From<u8> for OutlineAt {
    fn from(value: u8) -> Self {
        Self::Level(i64::from(value))
    }
}

impl From<u32> for OutlineAt {
    fn from(value: u32) -> Self {
        Self::Level(i64::from(value))
    }
}

impl From<i32> for OutlineAt {
    fn from(value: i32) -> Self {
        Self::Level(i64::from(value))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn outline_is_const_checked_for_rows_and_columns() {
        const VALID: Result<Outline, OutlineError> = OutlineAt::Level(7).resolve();
        const INVALID: Result<Outline, OutlineError> = OutlineAt::Level(8).resolve();
        assert_eq!(VALID.map(Outline::get), Ok(7));
        assert_eq!(INVALID.unwrap_err().value(), 8);
    }
}
