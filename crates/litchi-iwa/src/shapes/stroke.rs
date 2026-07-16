//! Typed native shape strokes and copy-on-write style CRUD.

use crate::{Error, Result};

use super::LineEndpoints;
#[cfg(test)]
use super::color::RgbColorSpace;
use super::color::RgbaColor;
const DEFAULT_MITER_LIMIT: f32 = 4.0;

mod native;
mod style;

#[cfg(test)]
use native::{pattern_to_native, stroke_from_native, stroke_to_native};
pub(crate) use style::{reset_shape_stroke, set_shape_stroke, shape_stroke};

/// A finite, positive stroke width measured in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct StrokeWidth(f32);

impl StrokeWidth {
    pub fn new(points: f32) -> Result<Self> {
        if !points.is_finite() || points <= 0.0 {
            return Err(Error::ParseError(
                "iWork stroke width must be finite and greater than zero".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    pub fn points(self) -> f32 {
        self.0
    }
}

/// Native standard stroke patterns exposed by Pages, Numbers, and Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum StrokePattern {
    #[default]
    Solid,
    ShortDash,
    MediumDash,
    LongDash,
    RoundedDash,
}

/// Geometry used at the ends of individual stroke segments.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum StrokeCap {
    #[default]
    Butt,
    Round,
    Square,
}

/// Geometry used where two stroke segments meet.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum StrokeJoin {
    #[default]
    Miter,
    Round,
    Bevel,
}

/// A finite, positive miter-limit ratio.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct StrokeMiterLimit(f32);

impl StrokeMiterLimit {
    pub fn new(ratio: f32) -> Result<Self> {
        if !ratio.is_finite() || ratio <= 0.0 {
            return Err(Error::ParseError(
                "iWork stroke miter limit must be finite and greater than zero".to_owned(),
            ));
        }
        Ok(Self(ratio))
    }

    pub fn ratio(self) -> f32 {
        self.0
    }
}

/// Fully typed native stroke appearance for an ordinary drawing shape.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeStroke {
    pub color: RgbaColor,
    pub width: StrokeWidth,
    pub pattern: StrokePattern,
    pub cap: StrokeCap,
    pub join: StrokeJoin,
    pub miter_limit: StrokeMiterLimit,
}

/// Stroke and endpoint appearance used when creating a native straight line.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct LineStyle {
    pub stroke: ShapeStroke,
    pub endpoints: LineEndpoints,
}

impl LineStyle {
    pub const fn new(stroke: ShapeStroke) -> Self {
        Self {
            stroke,
            endpoints: LineEndpoints::new(super::LineEndpoint::None, super::LineEndpoint::None),
        }
    }

    pub const fn with_endpoints(mut self, endpoints: LineEndpoints) -> Self {
        self.endpoints = endpoints;
        self
    }
}

impl ShapeStroke {
    /// Construct a standard app-native stroke. Rounded dash selects round cap/join.
    pub fn new(color: RgbaColor, width: StrokeWidth, pattern: StrokePattern) -> Self {
        let rounded = pattern == StrokePattern::RoundedDash;
        Self {
            color,
            width,
            pattern,
            cap: if rounded {
                StrokeCap::Round
            } else {
                StrokeCap::Butt
            },
            join: if rounded {
                StrokeJoin::Round
            } else {
                StrokeJoin::Miter
            },
            miter_limit: StrokeMiterLimit(DEFAULT_MITER_LIMIT),
        }
    }

    pub fn with_cap(mut self, cap: StrokeCap) -> Self {
        self.cap = cap;
        self
    }

    pub fn with_join(mut self, join: StrokeJoin) -> Self {
        self.join = join;
        self
    }

    pub fn with_miter_limit(mut self, miter_limit: StrokeMiterLimit) -> Self {
        self.miter_limit = miter_limit;
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn standard_patterns_match_real_app_archives() {
        for (pattern, expected) in [
            (StrokePattern::Solid, &[][..]),
            (StrokePattern::ShortDash, &[1.0, 1.0][..]),
            (StrokePattern::MediumDash, &[2.0, 2.0][..]),
            (StrokePattern::LongDash, &[6.0, 6.0][..]),
            (StrokePattern::RoundedDash, &[0.001, 2.0][..]),
        ] {
            let native = pattern_to_native(pattern);
            let count = native.count.unwrap() as usize;
            assert_eq!(&native.pattern[..count], expected);
        }
    }

    #[test]
    fn typed_stroke_round_trips_through_native_archive() {
        let stroke = ShapeStroke::new(
            RgbaColor::new(0.1, 0.2, 0.3, 0.8, RgbColorSpace::DisplayP3).unwrap(),
            StrokeWidth::new(3.5).unwrap(),
            StrokePattern::RoundedDash,
        );
        assert_eq!(
            stroke_from_native(&stroke_to_native(stroke)).unwrap(),
            Some(stroke)
        );
    }

    #[test]
    fn invalid_scalar_values_are_rejected() {
        assert!(StrokeWidth::new(0.0).is_err());
        assert!(StrokeWidth::new(f32::NAN).is_err());
        assert!(StrokeMiterLimit::new(f32::INFINITY).is_err());
        assert!(RgbaColor::new(-0.1, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
    }
}
