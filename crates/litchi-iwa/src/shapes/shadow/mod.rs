//! Typed drop, contact, and curved shadows for ordinary shapes.

mod native;
mod style;

use crate::{Error, Result};

use super::RgbaColor;

pub(crate) use style::{reset_shape_shadow, set_shape_shadow, shape_shadow};

const FULL_TURN_DEGREES: f32 = 360.0;
const RIGHT_ANGLE_DEGREES: f32 = 90.0;
const MAX_NATIVE_BLUR_RADIUS: u32 = i32::MAX as u32;

/// Canonical inspector angle in degrees, in the range `0.0..360.0`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeShadowAngle(f32);

impl ShapeShadowAngle {
    pub const ZERO: Self = Self(0.0);

    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() || !(0.0..FULL_TURN_DEGREES).contains(&degrees) {
            return Err(Error::ParseError(
                "iWork shadow angle must be finite and within 0.0..360.0 degrees".to_owned(),
            ));
        }
        Ok(Self(degrees))
    }

    pub const fn degrees(self) -> f32 {
        self.0
    }
}

/// Non-negative shadow offset measured in points.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeShadowOffset(f32);

impl ShapeShadowOffset {
    pub const ZERO: Self = Self(0.0);

    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() || points < 0.0 {
            return Err(Error::ParseError(
                "iWork shadow offset must be finite and non-negative".to_owned(),
            ));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Non-negative integral shadow blur radius measured in points.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct ShapeShadowBlurRadius(u32);

impl ShapeShadowBlurRadius {
    pub const ZERO: Self = Self(0);

    pub fn from_points(points: u32) -> Result<Self> {
        if points > MAX_NATIVE_BLUR_RADIUS {
            return Err(Error::ParseError(format!(
                "iWork shadow blur radius cannot exceed {MAX_NATIVE_BLUR_RADIUS} points"
            )));
        }
        Ok(Self(points))
    }

    pub const fn points(self) -> u32 {
        self.0
    }
}

/// Normalized shadow opacity in the inclusive range `0.0..=1.0`.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeShadowOpacity(f32);

impl ShapeShadowOpacity {
    pub const TRANSPARENT: Self = Self(0.0);
    pub const OPAQUE: Self = Self(1.0);

    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(value, "iWork shadow opacity")?;
        Ok(Self(value))
    }

    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for ShapeShadowOpacity {
    fn default() -> Self {
        Self::OPAQUE
    }
}

/// Contact-shadow viewing angle, stored losslessly as native height.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeShadowPerspective(f32);

impl ShapeShadowPerspective {
    pub const LEVEL: Self = Self(0.0);
    pub const MAXIMUM: Self = Self(1.0);

    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() || !(0.0..=RIGHT_ANGLE_DEGREES).contains(&degrees) {
            return Err(Error::ParseError(
                "iWork contact-shadow perspective must be finite and within 0..=90 degrees"
                    .to_owned(),
            ));
        }
        Ok(Self(degrees.to_radians().sin()))
    }

    pub fn degrees(self) -> f32 {
        self.0.asin().to_degrees()
    }

    pub(crate) fn from_native_height(height: f32) -> Result<Self> {
        validate_normalized(height, "iWork contact-shadow height")?;
        Ok(Self(height))
    }

    pub(crate) const fn native_height(self) -> f32 {
        self.0
    }
}

/// Signed curved-shadow bend: negative is inward and positive is outward.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeShadowCurve(f32);

impl ShapeShadowCurve {
    pub const FULLY_INWARD: Self = Self(-1.0);
    pub const FLAT: Self = Self(0.0);
    pub const FULLY_OUTWARD: Self = Self(1.0);

    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() || !(-1.0..=1.0).contains(&value) {
            return Err(Error::ParseError(
                "iWork curved-shadow bend must be finite and within -1.0..=1.0".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    pub const fn get(self) -> f32 {
        self.0
    }
}

/// Properties shared by every enabled iWork shadow family.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeShadowAppearance {
    color: RgbaColor,
    blur_radius: ShapeShadowBlurRadius,
    offset: ShapeShadowOffset,
    opacity: ShapeShadowOpacity,
}

impl ShapeShadowAppearance {
    pub const fn new(
        color: RgbaColor,
        blur_radius: ShapeShadowBlurRadius,
        offset: ShapeShadowOffset,
        opacity: ShapeShadowOpacity,
    ) -> Self {
        Self {
            color,
            blur_radius,
            offset,
            opacity,
        }
    }

    pub const fn color(self) -> RgbaColor {
        self.color
    }

    pub const fn blur_radius(self) -> ShapeShadowBlurRadius {
        self.blur_radius
    }

    pub const fn offset(self) -> ShapeShadowOffset {
        self.offset
    }

    pub const fn opacity(self) -> ShapeShadowOpacity {
        self.opacity
    }

    pub const fn with_color(mut self, color: RgbaColor) -> Self {
        self.color = color;
        self
    }

    pub const fn with_blur_radius(mut self, blur_radius: ShapeShadowBlurRadius) -> Self {
        self.blur_radius = blur_radius;
        self
    }

    pub const fn with_offset(mut self, offset: ShapeShadowOffset) -> Self {
        self.offset = offset;
        self
    }

    pub const fn with_opacity(mut self, opacity: ShapeShadowOpacity) -> Self {
        self.opacity = opacity;
        self
    }
}

/// App-native drop shadow.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeDropShadow {
    appearance: ShapeShadowAppearance,
    angle: ShapeShadowAngle,
}

impl ShapeDropShadow {
    pub const fn new(appearance: ShapeShadowAppearance, angle: ShapeShadowAngle) -> Self {
        Self { appearance, angle }
    }

    pub const fn appearance(self) -> ShapeShadowAppearance {
        self.appearance
    }

    pub const fn angle(self) -> ShapeShadowAngle {
        self.angle
    }
}

/// App-native contact shadow.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeContactShadow {
    appearance: ShapeShadowAppearance,
    perspective: ShapeShadowPerspective,
    contact_offset: ShapeShadowOffset,
}

impl ShapeContactShadow {
    pub const fn new(
        appearance: ShapeShadowAppearance,
        perspective: ShapeShadowPerspective,
    ) -> Self {
        Self {
            appearance,
            perspective,
            contact_offset: ShapeShadowOffset::ZERO,
        }
    }

    pub const fn appearance(self) -> ShapeShadowAppearance {
        self.appearance
    }

    pub const fn perspective(self) -> ShapeShadowPerspective {
        self.perspective
    }

    /// Native contact-plane offset. Current app inspectors normally leave this at zero.
    pub const fn contact_offset(self) -> ShapeShadowOffset {
        self.contact_offset
    }

    pub const fn with_contact_offset(mut self, offset: ShapeShadowOffset) -> Self {
        self.contact_offset = offset;
        self
    }
}

/// App-native curved shadow.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeCurvedShadow {
    appearance: ShapeShadowAppearance,
    angle: ShapeShadowAngle,
    curve: ShapeShadowCurve,
}

impl ShapeCurvedShadow {
    pub const fn new(
        appearance: ShapeShadowAppearance,
        angle: ShapeShadowAngle,
        curve: ShapeShadowCurve,
    ) -> Self {
        Self {
            appearance,
            angle,
            curve,
        }
    }

    pub const fn appearance(self) -> ShapeShadowAppearance {
        self.appearance
    }

    pub const fn angle(self) -> ShapeShadowAngle {
        self.angle
    }

    pub const fn curve(self) -> ShapeShadowCurve {
        self.curve
    }
}

/// Shadow state shown by the iWork Style inspector.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ShapeShadow {
    #[default]
    Disabled,
    Drop(ShapeDropShadow),
    Contact(ShapeContactShadow),
    Curved(ShapeCurvedShadow),
}

fn validate_normalized(value: f32, label: &str) -> Result<()> {
    if !value.is_finite() || !(0.0..=1.0).contains(&value) {
        return Err(Error::ParseError(format!(
            "{label} must be finite and within 0.0..=1.0"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shadow_scalars_reject_invalid_values() {
        assert!(ShapeShadowAngle::from_degrees(360.0).is_err());
        assert!(ShapeShadowAngle::from_degrees(f32::NAN).is_err());
        assert!(ShapeShadowOffset::from_points(-0.1).is_err());
        assert!(ShapeShadowOpacity::new(f32::INFINITY).is_err());
        assert!(ShapeShadowPerspective::from_degrees(90.1).is_err());
        assert!(ShapeShadowCurve::new(-1.01).is_err());
        assert!(ShapeShadowBlurRadius::from_points(MAX_NATIVE_BLUR_RADIUS + 1).is_err());
    }

    #[test]
    fn contact_perspective_preserves_native_height() {
        let perspective = ShapeShadowPerspective::from_degrees(23.0).unwrap();
        assert!((perspective.degrees() - 23.0).abs() < 0.000_1);
        assert_eq!(
            ShapeShadowPerspective::from_native_height(perspective.native_height()).unwrap(),
            perspective
        );
    }
}
