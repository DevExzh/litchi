//! Archive-free shape-stroke values shared by concrete iWork owners.

use crate::color::Rgba;
use crate::shape::line::{Endpoint, Endpoints};

const DEFAULT_MITER_LIMIT: f32 = 4.0;

/// Validation failures for shape-stroke values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A stroke width was not finite.
    #[error("shape stroke width must be finite")]
    WidthNonFinite,
    /// A stroke width was not positive.
    #[error("shape stroke width must be greater than zero")]
    WidthOutOfRange,
    /// A miter limit was not finite.
    #[error("shape stroke miter limit must be finite")]
    MiterLimitNonFinite,
    /// A miter limit was not positive.
    #[error("shape stroke miter limit must be greater than zero")]
    MiterLimitOutOfRange,
}

/// Result type for shape-stroke value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A finite, positive stroke width measured in points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Width(f32);

impl Width {
    /// One-point stroke used by standard iWork text outlines.
    pub const ONE: Self = Self(1.0);

    /// Construct a checked point width.
    ///
    /// # Errors
    ///
    /// Returns an error when `points` is non-finite or not positive.
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::WidthNonFinite);
        }
        if points <= 0.0 {
            return Err(Error::WidthOutOfRange);
        }
        Ok(Self(points))
    }

    /// Return the width in points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Native standard stroke patterns exposed by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Pattern {
    #[default]
    Solid,
    ShortDash,
    MediumDash,
    LongDash,
    RoundedDash,
}

/// Geometry used at the ends of individual stroke segments.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Cap {
    #[default]
    Butt,
    Round,
    Square,
}

/// Geometry used where two stroke segments meet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Join {
    #[default]
    Miter,
    Round,
    Bevel,
}

/// A finite, positive miter-limit ratio.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct MiterLimit(f32);

impl MiterLimit {
    /// Construct a checked miter-limit ratio.
    ///
    /// # Errors
    ///
    /// Returns an error when `ratio` is non-finite or not positive.
    pub fn new(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::MiterLimitNonFinite);
        }
        if ratio <= 0.0 {
            return Err(Error::MiterLimitOutOfRange);
        }
        Ok(Self(ratio))
    }

    /// Return the miter-limit ratio.
    #[must_use]
    pub const fn ratio(self) -> f32 {
        self.0
    }
}

/// Fully typed stroke appearance for an ordinary drawing shape.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Stroke {
    pub color: Rgba,
    pub width: Width,
    pub pattern: Pattern,
    pub cap: Cap,
    pub join: Join,
    pub miter_limit: MiterLimit,
}

impl Stroke {
    /// Construct a standard app-native stroke. Rounded dash selects round cap/join.
    #[must_use]
    pub fn new(color: Rgba, width: Width, pattern: Pattern) -> Self {
        let rounded = pattern == Pattern::RoundedDash;
        Self {
            color,
            width,
            pattern,
            cap: if rounded { Cap::Round } else { Cap::Butt },
            join: if rounded { Join::Round } else { Join::Miter },
            miter_limit: MiterLimit(DEFAULT_MITER_LIMIT),
        }
    }

    /// Set the segment cap.
    #[must_use]
    pub const fn with_cap(mut self, cap: Cap) -> Self {
        self.cap = cap;
        self
    }

    /// Set the segment join.
    #[must_use]
    pub const fn with_join(mut self, join: Join) -> Self {
        self.join = join;
        self
    }

    /// Set the miter-limit ratio.
    #[must_use]
    pub const fn with_miter_limit(mut self, miter_limit: MiterLimit) -> Self {
        self.miter_limit = miter_limit;
        self
    }
}

/// Stroke and endpoint appearance used when creating a native straight line.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct LineStyle {
    pub stroke: Stroke,
    pub endpoints: Endpoints,
}

impl LineStyle {
    /// Construct a line style without endpoint decorations.
    #[must_use]
    pub const fn new(stroke: Stroke) -> Self {
        Self {
            stroke,
            endpoints: Endpoints::new(Endpoint::None, Endpoint::None),
        }
    }

    /// Set the line endpoint decorations.
    #[must_use]
    pub const fn with_endpoints(mut self, endpoints: Endpoints) -> Self {
        self.endpoints = endpoints;
        self
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;
    use crate::color::RgbColorSpace;

    #[test]
    fn values_are_compact() {
        assert_eq!(size_of::<Width>(), 4);
        assert_eq!(size_of::<MiterLimit>(), 4);
        assert_eq!(size_of::<Stroke>(), 32);
    }

    #[test]
    fn scalar_values_reject_invalid_inputs() {
        assert!(Width::new(0.0).is_err());
        assert!(Width::new(f32::NAN).is_err());
        assert!(MiterLimit::new(f32::INFINITY).is_err());
        assert!(MiterLimit::new(-1.0).is_err());
    }

    #[test]
    fn rounded_dash_selects_rounded_geometry() {
        let stroke = Stroke::new(
            Rgba::new(0.1, 0.2, 0.3, 0.8, RgbColorSpace::DisplayP3).unwrap(),
            Width::new(3.5).unwrap(),
            Pattern::RoundedDash,
        );
        assert_eq!(stroke.cap, Cap::Round);
        assert_eq!(stroke.join, Join::Round);
    }
}
