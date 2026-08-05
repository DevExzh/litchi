//! Archive-free semantic values for native iWork shape paths.

/// Validation failures for shape-path controls.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The corner radius was not finite.
    #[error("corner radius must be finite")]
    RadiusNonFinite,
    /// The corner radius was negative.
    #[error("corner radius must be non-negative")]
    RadiusNegative,
    /// The polygon side count was below the native minimum.
    #[error("polygon side count must be at least 3")]
    PolygonSidesTooSmall,
    /// The star point count was below the native minimum.
    #[error("star point count must be at least 3")]
    StarPointsTooSmall,
    /// The inner-radius ratio was not finite.
    #[error("star inner-radius ratio must be finite")]
    InnerRadiusNonFinite,
    /// The inner-radius ratio was outside its half-open domain.
    #[error("star inner-radius ratio must be in [0, 1)")]
    InnerRadiusOutOfRange,
}

/// Result type for shape-path value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Corner radius in the path's natural coordinate system.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct CornerRadius(f32);

impl CornerRadius {
    /// Native default used by iWork for a 100-point rounded rectangle.
    pub const DEFAULT: Self = Self(15.0);

    /// Validate a corner radius measured in path points.
    ///
    /// # Errors
    ///
    /// Returns [`Error::RadiusNonFinite`] for NaN or infinity and
    /// [`Error::RadiusNegative`] for a negative radius.
    #[must_use = "use the validated radius or handle its validation error"]
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::RadiusNonFinite);
        }
        if points < 0.0 {
            return Err(Error::RadiusNegative);
        }
        Ok(Self(points))
    }

    /// Return the radius in path points.
    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Valid number of sides for an iWork regular polygon.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct PolygonSides(u8);

impl PolygonSides {
    /// Three-sided regular polygon.
    pub const TRIANGLE: Self = Self(3);
    /// Five-sided regular polygon used by iWork's Pentagon preset.
    pub const PENTAGON: Self = Self(5);

    /// Validate a regular-polygon side count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PolygonSidesTooSmall`] when `sides` is below three.
    #[must_use = "use the validated side count or handle its validation error"]
    pub fn new(sides: u8) -> Result<Self> {
        if sides < 3 {
            return Err(Error::PolygonSidesTooSmall);
        }
        Ok(Self(sides))
    }

    /// Return the side count.
    #[must_use]
    pub const fn count(self) -> u8 {
        self.0
    }
}

/// Valid number of outer points for an iWork star.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct StarPoints(u8);

impl StarPoints {
    /// Native five-point star control.
    pub const FIVE: Self = Self(5);

    /// Validate a star point count.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StarPointsTooSmall`] when `points` is below three.
    #[must_use = "use the validated point count or handle its validation error"]
    pub fn new(points: u8) -> Result<Self> {
        if points < 3 {
            return Err(Error::StarPointsTooSmall);
        }
        Ok(Self(points))
    }

    /// Return the outer point count.
    #[must_use]
    pub const fn count(self) -> u8 {
        self.0
    }
}

/// Ratio between an iWork star's inner and outer radii.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct InnerRadiusRatio(f32);

impl InnerRadiusRatio {
    /// Native default used by iWork's five-point star preset.
    pub const DEFAULT: Self = Self(0.382);

    /// Validate an inner-radius ratio in the half-open interval `[0, 1)`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InnerRadiusNonFinite`] for NaN or infinity and
    /// [`Error::InnerRadiusOutOfRange`] outside `[0, 1)`.
    #[must_use = "use the validated ratio or handle its validation error"]
    pub fn new(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() {
            return Err(Error::InnerRadiusNonFinite);
        }
        if !(0.0..1.0).contains(&ratio) {
            return Err(Error::InnerRadiusOutOfRange);
        }
        Ok(Self(ratio))
    }

    /// Return the inner-radius ratio.
    #[must_use]
    pub const fn ratio(self) -> f32 {
        self.0
    }
}

/// Source-buildable iWork shape preset.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Preset {
    /// Four-corner Bézier rectangle.
    Rectangle,
    /// Native rounded rectangle with an explicit corner radius.
    RoundedRectangle { corner_radius: CornerRadius },
    /// Four-segment cubic Bézier ellipse.
    Ellipse,
    /// Native left-facing single arrow with standard head and shaft proportions.
    LeftArrow,
    /// Native right-facing single arrow with standard head and shaft proportions.
    RightArrow,
    /// Native bidirectional arrow with standard head and shaft proportions.
    DoubleArrow,
    /// Native regular polygon with a configurable side count.
    RegularPolygon { sides: PolygonSides },
    /// Native star with configurable point count and inner radius.
    Star {
        points: StarPoints,
        inner_radius: InnerRadiusRatio,
    },
}

impl Preset {
    /// Native default rounded rectangle.
    pub const ROUNDED_RECTANGLE: Self = Self::RoundedRectangle {
        corner_radius: CornerRadius::DEFAULT,
    };
    /// Native five-sided Pentagon preset.
    pub const PENTAGON: Self = Self::RegularPolygon {
        sides: PolygonSides::PENTAGON,
    };
    /// Native five-point Star preset.
    pub const STAR: Self = Self::Star {
        points: StarPoints::FIVE,
        inner_radius: InnerRadiusRatio::DEFAULT,
    };
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{CornerRadius, Error, InnerRadiusRatio, PolygonSides, Preset, StarPoints};

    #[test]
    fn scalar_controls_are_compact_and_copyable() {
        assert_eq!(size_of::<CornerRadius>(), 4);
        assert_eq!(size_of::<PolygonSides>(), 1);
        assert_eq!(size_of::<StarPoints>(), 1);
        assert_eq!(size_of::<InnerRadiusRatio>(), 4);
        assert_eq!(align_of::<CornerRadius>(), 4);
        assert_eq!(align_of::<PolygonSides>(), 1);
        assert_eq!(align_of::<StarPoints>(), 1);
        assert_eq!(align_of::<InnerRadiusRatio>(), 4);

        let preset = Preset::STAR;
        let copied = preset;
        assert_eq!(preset, copied);
    }

    #[test]
    fn controls_reject_non_domain_values() {
        assert_eq!(CornerRadius::new(f32::NAN), Err(Error::RadiusNonFinite));
        assert_eq!(
            CornerRadius::new(f32::INFINITY),
            Err(Error::RadiusNonFinite)
        );
        assert_eq!(CornerRadius::new(-0.01), Err(Error::RadiusNegative));
        assert_eq!(PolygonSides::new(2), Err(Error::PolygonSidesTooSmall));
        assert_eq!(StarPoints::new(2), Err(Error::StarPointsTooSmall));
        assert_eq!(
            InnerRadiusRatio::new(f32::NAN),
            Err(Error::InnerRadiusNonFinite)
        );
        assert_eq!(
            InnerRadiusRatio::new(-0.01),
            Err(Error::InnerRadiusOutOfRange)
        );
        assert_eq!(
            InnerRadiusRatio::new(1.0),
            Err(Error::InnerRadiusOutOfRange)
        );
    }

    #[test]
    fn valid_controls_preserve_native_values() {
        assert_eq!(
            CornerRadius::new(-0.0).unwrap().points().to_bits(),
            (-0.0f32).to_bits()
        );
        assert_eq!(PolygonSides::new(8).unwrap().count(), 8);
        assert_eq!(StarPoints::new(7).unwrap().count(), 7);
        assert_eq!(
            InnerRadiusRatio::new(0.45).unwrap().ratio().to_bits(),
            0.45f32.to_bits()
        );
    }
}
