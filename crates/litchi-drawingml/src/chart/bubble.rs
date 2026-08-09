//! Bubble-chart-specific value types.

use std::str::FromStr;

use thiserror::Error;

/// How a bubble's numeric size is interpreted.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Size {
    /// The value controls the bubble's area.
    #[default]
    Area,
    /// The value controls the bubble's width.
    Width,
}

impl Size {
    #[inline]
    pub(crate) const fn xml_value(self) -> &'static str {
        match self {
            Self::Area => "area",
            Self::Width => "w",
        }
    }

    #[inline]
    pub(crate) fn from_xml(value: &[u8]) -> Result<Self, SizeError> {
        match value {
            b"area" => Ok(Self::Area),
            b"w" => Ok(Self::Width),
            _ => Err(SizeError),
        }
    }
}

impl FromStr for Size {
    type Err = SizeError;

    #[inline]
    fn from_str(value: &str) -> Result<Self, Self::Err> {
        Self::from_xml(value.as_bytes())
    }
}

/// An invalid bubble-size representation token.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[error("bubble size must be represented by area ('area') or width ('w')")]
pub struct SizeError;

/// A bubble scale percentage in the inclusive range 0–300.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Scale(u16);

impl Scale {
    /// Smallest supported scale percentage.
    pub const MIN: u16 = 0;
    /// Largest supported scale percentage.
    pub const MAX: u16 = 300;
    /// Schema default scale percentage.
    pub const DEFAULT: Self = Self(100);

    /// Creates a checked scale percentage.
    #[inline]
    pub const fn new(value: u16) -> Result<Self, ScaleError> {
        if value <= Self::MAX {
            Ok(Self(value))
        } else {
            Err(ScaleError {
                value: value as u32,
            })
        }
    }

    /// Returns the scale percentage.
    #[inline]
    #[must_use]
    pub const fn get(self) -> u16 {
        self.0
    }
}

impl Default for Scale {
    #[inline]
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl TryFrom<u32> for Scale {
    type Error = ScaleError;

    #[inline]
    fn try_from(value: u32) -> Result<Self, Self::Error> {
        u16::try_from(value)
            .ok()
            .and_then(|value| Self::new(value).ok())
            .ok_or(ScaleError { value })
    }
}

impl From<Scale> for u16 {
    #[inline]
    fn from(value: Scale) -> Self {
        value.get()
    }
}

impl From<Scale> for u32 {
    #[inline]
    fn from(value: Scale) -> Self {
        u32::from(value.get())
    }
}

/// A value outside the supported bubble-scale range.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[error("bubble scale {value} is outside the inclusive range 0..=300")]
pub struct ScaleError {
    value: u32,
}

impl ScaleError {
    /// Returns the rejected scale percentage.
    #[inline]
    #[must_use]
    pub const fn value(self) -> u32 {
        self.value
    }
}

#[cfg(test)]
mod tests {
    use super::{Scale, Size};

    #[test]
    fn scale_accepts_boundaries_and_rejects_values_above_the_range() {
        assert_eq!(Scale::new(Scale::MIN).map(Scale::get), Ok(0));
        assert_eq!(Scale::new(Scale::MAX).map(Scale::get), Ok(300));
        assert_eq!(Scale::new(301).unwrap_err().value(), 301);
        assert_eq!(Scale::try_from(u32::MAX).unwrap_err().value(), u32::MAX);
    }

    #[test]
    fn size_tokens_round_trip_without_allocation() {
        for size in [Size::Area, Size::Width] {
            let token = size.xml_value();
            assert_eq!(token.parse::<Size>(), Ok(size));
        }
        assert!("diameter".parse::<Size>().is_err());
    }
}
