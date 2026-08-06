//! Validated bounded scalar values.

use serde::{Deserialize, Deserializer, Serialize, Serializer, de::Error as _};
use thiserror::Error;

/// A `u32` proven to be inside the inclusive const-generic range.
///
/// Construction is checked and the inner value is private, so safe code cannot
/// create an out-of-range value. `new` is `const`, allowing literal validation
/// during constant evaluation without a panicking constructor.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[allow(
    clippy::module_name_repetitions,
    reason = "public API name is stable and used by dependent crates; renaming would be a breaking change"
)]
pub struct BoundedU32<const MIN: u32, const MAX: u32>(u32);

impl<const MIN: u32, const MAX: u32> BoundedU32<MIN, MAX> {
    /// Creates a bounded value, returning `None` for an invalid range or value.
    #[inline]
    #[must_use]
    pub const fn new(value: u32) -> Option<Self> {
        if MIN <= MAX && value >= MIN && value <= MAX {
            Some(Self(value))
        } else {
            None
        }
    }

    /// Creates a bounded value with a structured error.
    ///
    /// # Errors
    ///
    /// Returns `BoundsError::InvalidRange` if `MIN > MAX`, or
    /// `BoundsError::OutOfRange` if `value` is outside `MIN..=MAX`.
    #[inline]
    pub const fn try_new(value: u32) -> Result<Self, BoundsError> {
        if MIN > MAX {
            return Err(BoundsError::InvalidRange { min: MIN, max: MAX });
        }
        if value < MIN || value > MAX {
            return Err(BoundsError::OutOfRange {
                value,
                min: MIN,
                max: MAX,
            });
        }
        Ok(Self(value))
    }

    /// Returns the validated value.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }

    /// Inclusive lower bound.
    pub const MIN: u32 = MIN;

    /// Inclusive upper bound.
    pub const MAX: u32 = MAX;
}

impl<const MIN: u32, const MAX: u32> std::fmt::Debug for BoundedU32<MIN, MAX> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_tuple("BoundedU32").field(&self.0).finish()
    }
}

impl<const MIN: u32, const MAX: u32> std::fmt::Display for BoundedU32<MIN, MAX> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

impl<const MIN: u32, const MAX: u32> TryFrom<u32> for BoundedU32<MIN, MAX> {
    type Error = BoundsError;

    fn try_from(value: u32) -> Result<Self, Self::Error> {
        Self::try_new(value)
    }
}

impl<const MIN: u32, const MAX: u32> From<BoundedU32<MIN, MAX>> for u32 {
    fn from(value: BoundedU32<MIN, MAX>) -> Self {
        value.get()
    }
}

impl<const MIN: u32, const MAX: u32> Serialize for BoundedU32<MIN, MAX> {
    fn serialize<S>(&self, serializer: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer,
    {
        serializer.serialize_u32(self.0)
    }
}

impl<'de, const MIN: u32, const MAX: u32> Deserialize<'de> for BoundedU32<MIN, MAX> {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>,
    {
        let value = u32::deserialize(deserializer)?;
        Self::try_new(value).map_err(D::Error::custom)
    }
}

/// Failure to construct a bounded scalar.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BoundsError {
    /// The type itself declares an empty range.
    #[error("invalid inclusive range {min}..={max}")]
    InvalidRange { min: u32, max: u32 },
    /// The value lies outside the declared inclusive range.
    #[error("value {value} is outside inclusive range {min}..={max}")]
    OutOfRange { value: u32, min: u32, max: u32 },
}

#[cfg(test)]
mod tests {
    use super::*;

    const HALF: Option<Percent> = Percent::new(50);
    const INVALID: Option<Percent> = Percent::new(101);

    type Percent = BoundedU32<0, 100>;

    #[test]
    fn validates_in_const_context_without_panicking() {
        assert_eq!(HALF.map(BoundedU32::get), Some(50));
        assert_eq!(INVALID, None);
    }

    #[test]
    fn reports_invalid_type_ranges() {
        type Empty = BoundedU32<2, 1>;
        assert_eq!(
            Empty::try_new(1),
            Err(BoundsError::InvalidRange { min: 2, max: 1 })
        );
    }
}
