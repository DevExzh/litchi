//! Archive-free numeric bounds for a native chart value axis.

/// Validation failures for value-axis bounds.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A bound was not finite.
    #[error("chart axis bound must be finite")]
    NonFinite,
    /// The lower bound is greater than the upper bound.
    #[error("chart value-axis minimum exceeds maximum")]
    Inverted,
}

/// Result type for value-axis bound construction.
pub type Result<T> = std::result::Result<T, Error>;

/// One finite manual bound for a chart value axis.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Bound(f64);

impl Bound {
    /// Create a finite chart-axis bound.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinity.
    #[must_use = "use the validated bound or handle its validation error"]
    pub fn new(value: f64) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::NonFinite);
        }
        Ok(Self(value))
    }

    /// Return the value represented in the native inspector.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.0
    }
}

impl TryFrom<f64> for Bound {
    type Error = Error;

    fn try_from(value: f64) -> Result<Self> {
        Self::new(value)
    }
}

/// Optional manual bounds for a chart's primary value axis.
///
/// A missing endpoint is the native automatic value. The constructor rejects
/// inverted ranges before an archive adapter can mutate a package.
#[repr(C)]
#[derive(Debug, Clone, Copy)]
pub struct Bounds {
    minimum: f64,
    maximum: f64,
    present: u8,
}

impl Bounds {
    /// Build a value-axis range from optional manual endpoints.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Inverted`] when both endpoints are present and the
    /// minimum is greater than the maximum.
    #[must_use = "use the validated bounds or handle their validation error"]
    pub fn new(minimum: Option<Bound>, maximum: Option<Bound>) -> Result<Self> {
        let bounds = Self {
            minimum: minimum.map_or(0.0, Bound::value),
            maximum: maximum.map_or(0.0, Bound::value),
            present: u8::from(minimum.is_some()) | (u8::from(maximum.is_some()) << 1),
        };
        bounds.validate()?;
        Ok(bounds)
    }

    /// Use native automatic lower and upper bounds.
    #[must_use]
    pub const fn automatic() -> Self {
        Self {
            minimum: 0.0,
            maximum: 0.0,
            present: 0,
        }
    }

    /// Build a fully manual value-axis range.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Inverted`] when `minimum` is greater than `maximum`.
    #[must_use = "use the validated bounds or handle their validation error"]
    pub fn fixed(minimum: Bound, maximum: Bound) -> Result<Self> {
        Self::new(Some(minimum), Some(maximum))
    }

    /// Return the optional manual lower bound.
    #[must_use]
    pub const fn minimum(self) -> Option<Bound> {
        if self.present & 1 != 0 {
            Some(Bound(self.minimum))
        } else {
            None
        }
    }

    /// Return the optional manual upper bound.
    #[must_use]
    pub const fn maximum(self) -> Option<Bound> {
        if self.present & 2 != 0 {
            Some(Bound(self.maximum))
        } else {
            None
        }
    }

    fn validate(self) -> Result<()> {
        if let (Some(minimum), Some(maximum)) = (self.minimum(), self.maximum())
            && minimum.value() > maximum.value()
        {
            return Err(Error::Inverted);
        }
        Ok(())
    }
}

impl PartialEq for Bounds {
    fn eq(&self, other: &Self) -> bool {
        self.present == other.present
            && (self.present & 1 == 0 || self.minimum == other.minimum)
            && (self.present & 2 == 0 || self.maximum == other.maximum)
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{Bound, Bounds, Error};

    #[test]
    fn bounds_are_compact_and_strict() {
        assert_eq!(size_of::<Bound>(), 8);
        assert_eq!(align_of::<Bound>(), 8);
        assert_eq!(size_of::<Bounds>(), 24);
        assert_eq!(Bound::new(f64::NAN), Err(Error::NonFinite));
        assert_eq!(Bound::new(f64::INFINITY), Err(Error::NonFinite));
        let low = Bound::new(-1.0).unwrap();
        let high = Bound::new(1.0).unwrap();
        assert_eq!(Bounds::fixed(high, low), Err(Error::Inverted));
        assert_eq!(Bounds::fixed(low, high).unwrap().minimum(), Some(low));
    }
}
