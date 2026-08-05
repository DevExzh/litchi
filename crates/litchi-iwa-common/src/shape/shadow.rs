//! Archive-free shape-shadow values shared by concrete iWork owners.

use crate::color::Rgba;

const FULL_TURN_DEGREES: f32 = 360.0;
const RIGHT_ANGLE_DEGREES: f32 = 90.0;
const MAX_BLUR_RADIUS: u32 = i32::MAX as u32;

/// Validation failures for shape-shadow values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// An angle was not finite.
    #[error("shape shadow angle must be finite")]
    AngleNonFinite,
    /// An angle was outside `[0, 360)` degrees.
    #[error("shape shadow angle must be in [0, 360) degrees")]
    AngleOutOfRange,
    /// An offset was not finite.
    #[error("shape shadow offset must be finite")]
    OffsetNonFinite,
    /// An offset was negative.
    #[error("shape shadow offset must be non-negative")]
    OffsetOutOfRange,
    /// A blur radius was outside the native signed range.
    #[error("shape shadow blur radius exceeds the native maximum")]
    BlurRadiusOutOfRange,
    /// An opacity was not finite.
    #[error("shape shadow opacity must be finite")]
    OpacityNonFinite,
    /// An opacity was outside the normalized range.
    #[error("shape shadow opacity must be in 0.0..=1.0")]
    OpacityOutOfRange,
    /// A perspective was not finite.
    #[error("shape shadow perspective must be finite")]
    PerspectiveNonFinite,
    /// A perspective was outside the inclusive right-angle range.
    #[error("shape shadow perspective must be in 0..=90 degrees")]
    PerspectiveOutOfRange,
    /// A normalized perspective height was outside its range.
    #[error("shape shadow perspective height must be in 0.0..=1.0")]
    PerspectiveHeightOutOfRange,
    /// A curve was not finite.
    #[error("shape shadow curve must be finite")]
    CurveNonFinite,
    /// A curve was outside the normalized signed range.
    #[error("shape shadow curve must be in -1.0..=1.0")]
    CurveOutOfRange,
}

/// Result type for shape-shadow value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Inspector angle in degrees, in the range `0.0..360.0`.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Angle(f32);

impl Angle {
    pub const ZERO: Self = Self(0.0);
    /// Forty-five-degree direction used by standard iWork text shadows.
    pub const FORTY_FIVE_DEGREES: Self = Self(45.0);

    /// Construct an inspector angle in degrees.
    ///
    /// # Errors
    ///
    /// Returns an error when `degrees` is non-finite or outside `[0, 360)`.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() {
            return Err(Error::AngleNonFinite);
        }
        if !(0.0..FULL_TURN_DEGREES).contains(&degrees) {
            return Err(Error::AngleOutOfRange);
        }
        Ok(Self(degrees))
    }

    #[must_use]
    pub const fn degrees(self) -> f32 {
        self.0
    }
}

/// Non-negative shadow offset measured in points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Offset(f32);

impl Offset {
    pub const ZERO: Self = Self(0.0);
    /// Five-point offset used by standard iWork text shadows.
    pub const FIVE_POINTS: Self = Self(5.0);
    /// Six-point offset used by newly inserted native chart shadows.
    pub const SIX_POINTS: Self = Self(6.0);

    /// Construct a checked point offset.
    ///
    /// # Errors
    ///
    /// Returns an error when `points` is non-finite or negative.
    pub fn from_points(points: f32) -> Result<Self> {
        if !points.is_finite() {
            return Err(Error::OffsetNonFinite);
        }
        if points < 0.0 {
            return Err(Error::OffsetOutOfRange);
        }
        Ok(Self(points))
    }

    #[must_use]
    pub const fn points(self) -> f32 {
        self.0
    }
}

/// Non-negative integral shadow blur radius measured in points.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct BlurRadius(u32);

impl BlurRadius {
    pub const ZERO: Self = Self(0);
    /// One-point blur radius used by standard iWork text shadows.
    pub const ONE_POINT: Self = Self(1);
    /// Ten-point blur used by newly inserted native chart shadows.
    pub const TEN_POINTS: Self = Self(10);

    /// Construct a checked point radius.
    ///
    /// # Errors
    ///
    /// Returns an error when `points` exceeds the native signed maximum.
    pub fn from_points(points: u32) -> Result<Self> {
        if points > MAX_BLUR_RADIUS {
            return Err(Error::BlurRadiusOutOfRange);
        }
        Ok(Self(points))
    }

    #[must_use]
    pub const fn points(self) -> u32 {
        self.0
    }
}

/// Normalized shadow opacity in the inclusive range `0.0..=1.0`.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Opacity(f32);

impl Opacity {
    pub const TRANSPARENT: Self = Self(0.0);
    /// Seventy-five percent opacity used by newly inserted native chart shadows.
    pub const THREE_QUARTERS: Self = Self(0.75);
    pub const OPAQUE: Self = Self(1.0);

    /// Construct a checked normalized opacity.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite or outside `0.0..=1.0`.
    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(value, Error::OpacityNonFinite, Error::OpacityOutOfRange)?;
        Ok(Self(value))
    }

    #[must_use]
    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for Opacity {
    fn default() -> Self {
        Self::OPAQUE
    }
}

/// Contact-shadow viewing angle, stored as normalized native height.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Perspective(f32);

impl Perspective {
    pub const LEVEL: Self = Self(0.0);
    pub const MAXIMUM: Self = Self(1.0);

    /// Construct a perspective from the inspector's degree value.
    ///
    /// # Errors
    ///
    /// Returns an error when `degrees` is non-finite or outside `0..=90`.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        if !degrees.is_finite() {
            return Err(Error::PerspectiveNonFinite);
        }
        if !(0.0..=RIGHT_ANGLE_DEGREES).contains(&degrees) {
            return Err(Error::PerspectiveOutOfRange);
        }
        Ok(Self(degrees.to_radians().sin()))
    }

    /// Return the inspector's degree value.
    #[must_use]
    pub fn degrees(self) -> f32 {
        self.0.asin().to_degrees()
    }

    /// Construct from the normalized contact-plane height used by iWork.
    ///
    /// # Errors
    ///
    /// Returns an error when `height` is non-finite or outside `0.0..=1.0`.
    pub fn from_height(height: f32) -> Result<Self> {
        validate_normalized(
            height,
            Error::PerspectiveNonFinite,
            Error::PerspectiveHeightOutOfRange,
        )?;
        Ok(Self(height))
    }

    /// Return the normalized contact-plane height.
    #[must_use]
    pub const fn height(self) -> f32 {
        self.0
    }
}

/// Signed curved-shadow bend: negative is inward and positive is outward.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Curve(f32);

impl Curve {
    pub const FULLY_INWARD: Self = Self(-1.0);
    pub const FLAT: Self = Self(0.0);
    pub const FULLY_OUTWARD: Self = Self(1.0);

    /// Construct a checked normalized curve.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` is non-finite or outside `-1.0..=1.0`.
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::CurveNonFinite);
        }
        if !(-1.0..=1.0).contains(&value) {
            return Err(Error::CurveOutOfRange);
        }
        Ok(Self(value))
    }

    #[must_use]
    pub const fn get(self) -> f32 {
        self.0
    }
}

/// Properties shared by every enabled iWork shadow family.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Appearance {
    color: Rgba,
    blur_radius: BlurRadius,
    offset: Offset,
    opacity: Opacity,
}

impl Appearance {
    #[must_use]
    pub const fn new(
        color: Rgba,
        blur_radius: BlurRadius,
        offset: Offset,
        opacity: Opacity,
    ) -> Self {
        Self {
            color,
            blur_radius,
            offset,
            opacity,
        }
    }

    #[must_use]
    pub const fn color(self) -> Rgba {
        self.color
    }

    #[must_use]
    pub const fn blur_radius(self) -> BlurRadius {
        self.blur_radius
    }

    #[must_use]
    pub const fn offset(self) -> Offset {
        self.offset
    }

    #[must_use]
    pub const fn opacity(self) -> Opacity {
        self.opacity
    }

    #[must_use]
    pub const fn with_color(mut self, color: Rgba) -> Self {
        self.color = color;
        self
    }

    #[must_use]
    pub const fn with_blur_radius(mut self, blur_radius: BlurRadius) -> Self {
        self.blur_radius = blur_radius;
        self
    }

    #[must_use]
    pub const fn with_offset(mut self, offset: Offset) -> Self {
        self.offset = offset;
        self
    }

    #[must_use]
    pub const fn with_opacity(mut self, opacity: Opacity) -> Self {
        self.opacity = opacity;
        self
    }
}

/// App-native drop shadow.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Drop {
    appearance: Appearance,
    angle: Angle,
}

impl Drop {
    #[must_use]
    pub const fn new(appearance: Appearance, angle: Angle) -> Self {
        Self { appearance, angle }
    }

    #[must_use]
    pub const fn appearance(self) -> Appearance {
        self.appearance
    }

    #[must_use]
    pub const fn angle(self) -> Angle {
        self.angle
    }
}

/// App-native contact shadow.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Contact {
    appearance: Appearance,
    perspective: Perspective,
    offset: Offset,
}

impl Contact {
    #[must_use]
    pub const fn new(appearance: Appearance, perspective: Perspective) -> Self {
        Self {
            appearance,
            perspective,
            offset: Offset::ZERO,
        }
    }

    #[must_use]
    pub const fn appearance(self) -> Appearance {
        self.appearance
    }

    #[must_use]
    pub const fn perspective(self) -> Perspective {
        self.perspective
    }

    /// Native contact-plane offset. Current app inspectors normally leave this at zero.
    #[must_use]
    pub const fn contact_offset(self) -> Offset {
        self.offset
    }

    #[must_use]
    pub const fn with_contact_offset(mut self, offset: Offset) -> Self {
        self.offset = offset;
        self
    }
}

/// App-native curved shadow.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Curved {
    appearance: Appearance,
    angle: Angle,
    curve: Curve,
}

impl Curved {
    #[must_use]
    pub const fn new(appearance: Appearance, angle: Angle, curve: Curve) -> Self {
        Self {
            appearance,
            angle,
            curve,
        }
    }

    #[must_use]
    pub const fn appearance(self) -> Appearance {
        self.appearance
    }

    #[must_use]
    pub const fn angle(self) -> Angle {
        self.angle
    }

    #[must_use]
    pub const fn curve(self) -> Curve {
        self.curve
    }
}

/// Shadow state shown by the iWork Style inspector.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Shadow {
    #[default]
    Disabled,
    Drop(Drop),
    Contact(Contact),
    Curved(Curved),
}

fn validate_normalized(value: f32, non_finite: Error, out_of_range: Error) -> Result<()> {
    if !value.is_finite() {
        return Err(non_finite);
    }
    if !(0.0..=1.0).contains(&value) {
        return Err(out_of_range);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;
    use crate::color::{RgbColorSpace, Rgba};

    fn appearance(radius: u32, offset: f32, opacity: f32) -> Appearance {
        Appearance::new(
            Rgba::new(0.0, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).unwrap(),
            BlurRadius::from_points(radius).unwrap(),
            Offset::from_points(offset).unwrap(),
            Opacity::new(opacity).unwrap(),
        )
    }

    #[test]
    fn scalar_values_are_compact() {
        assert_eq!(size_of::<Angle>(), 4);
        assert_eq!(size_of::<Offset>(), 4);
        assert_eq!(size_of::<BlurRadius>(), 4);
        assert_eq!(size_of::<Opacity>(), 4);
        assert_eq!(size_of::<Perspective>(), 4);
        assert_eq!(size_of::<Curve>(), 4);
    }

    #[test]
    fn shadow_scalars_reject_invalid_values() {
        assert!(Angle::from_degrees(360.0).is_err());
        assert!(Angle::from_degrees(f32::NAN).is_err());
        assert!(Offset::from_points(-0.1).is_err());
        assert!(Opacity::new(f32::INFINITY).is_err());
        assert!(Perspective::from_degrees(90.1).is_err());
        assert!(Curve::new(-1.01).is_err());
        assert!(BlurRadius::from_points(MAX_BLUR_RADIUS + 1).is_err());
    }

    #[test]
    fn contact_perspective_preserves_normalized_height() {
        let perspective = Perspective::from_degrees(23.0).unwrap();
        assert!((perspective.degrees() - 23.0).abs() < 0.000_1);
        assert_eq!(
            Perspective::from_height(perspective.height()).unwrap(),
            perspective
        );
    }

    #[test]
    fn all_shadow_families_are_copyable_values() {
        let shadow = Shadow::Curved(Curved::new(
            appearance(7, 3.0, 0.5),
            Angle::from_degrees(135.0).unwrap(),
            Curve::new(0.2).unwrap(),
        ));
        assert_eq!(shadow, shadow);
    }
}
