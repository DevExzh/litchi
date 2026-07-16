//! Strict public model for native iWork shape gradients.

use std::f32::consts::TAU;

use crate::{Error, Result};

use super::super::RgbaColor;

const MINIMUM_STOP_COUNT: usize = 2;
const SIMPLE_STOP_COUNT: usize = 2;
const DEFAULT_MIDPOINT: f32 = 0.5;

/// Native geometry used to paint a shape gradient.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum ShapeGradientKind {
    /// Colors progress along a straight axis.
    Linear,
    /// Colors radiate from the native gradient origin.
    Radial,
}

/// Angle of a shape gradient, stored in native radians.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeGradientAngle(f32);

impl ShapeGradientAngle {
    /// Construct an angle from the degree value shown by iWork.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        Self::from_radians(degrees.to_radians())
    }

    /// Construct an angle from the native radian value.
    pub fn from_radians(radians: f32) -> Result<Self> {
        if !radians.is_finite() || !(0.0..TAU).contains(&radians) {
            return Err(Error::ParseError(
                "iWork shape gradient angle must be finite and in [0, 2π) radians".to_owned(),
            ));
        }
        Ok(Self(radians))
    }

    /// Return the exact radian value stored in the native archive.
    pub const fn radians(self) -> f32 {
        self.0
    }

    /// Return the angle in the degree unit displayed by iWork.
    pub fn degrees(self) -> f32 {
        self.0.to_degrees()
    }
}

macro_rules! unit_fraction {
    ($name:ident, $context:literal) => {
        #[doc = $context]
        #[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
        pub struct $name(f32);

        impl $name {
            /// Construct a finite normalized value in the inclusive range `0..=1`.
            pub fn new(value: f32) -> Result<Self> {
                if !value.is_finite() || !(0.0..=1.0).contains(&value) {
                    return Err(Error::ParseError(
                        concat!($context, " must be finite and between 0 and 1").to_owned(),
                    ));
                }
                Ok(Self(value))
            }

            /// Return the normalized value.
            pub const fn get(self) -> f32 {
                self.0
            }
        }
    };
}

unit_fraction!(
    ShapeGradientStopPosition,
    "Normalized location of a shape-gradient color stop."
);
unit_fraction!(
    ShapeGradientStopMidpoint,
    "Normalized blend midpoint following a shape-gradient color stop."
);
unit_fraction!(
    ShapeGradientOpacity,
    "Normalized overall opacity of a shape gradient."
);

impl ShapeGradientStopPosition {
    /// The beginning of a gradient axis.
    pub const START: Self = Self(0.0);
    /// The end of a gradient axis.
    pub const END: Self = Self(1.0);
}

impl ShapeGradientStopMidpoint {
    /// The midpoint used by iWork's simple two-color gradient editor.
    pub const CENTER: Self = Self(DEFAULT_MIDPOINT);
}

impl ShapeGradientOpacity {
    /// A fully opaque gradient.
    pub const OPAQUE: Self = Self(1.0);
}

/// One immutable, validated color stop in a native shape gradient.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ShapeGradientStop {
    color: RgbaColor,
    position: ShapeGradientStopPosition,
    midpoint: ShapeGradientStopMidpoint,
}

impl ShapeGradientStop {
    /// Construct one color stop from validated components.
    pub const fn new(
        color: RgbaColor,
        position: ShapeGradientStopPosition,
        midpoint: ShapeGradientStopMidpoint,
    ) -> Self {
        Self {
            color,
            position,
            midpoint,
        }
    }

    pub const fn color(self) -> RgbaColor {
        self.color
    }

    pub const fn position(self) -> ShapeGradientStopPosition {
        self.position
    }

    pub const fn midpoint(self) -> ShapeGradientStopMidpoint {
        self.midpoint
    }
}

/// Native simple or advanced linear/radial shape gradient.
#[derive(Debug, Clone, PartialEq)]
pub struct ShapeGradient {
    kind: ShapeGradientKind,
    stops: Box<[ShapeGradientStop]>,
    opacity: ShapeGradientOpacity,
    is_advanced: bool,
    angle: ShapeGradientAngle,
}

impl ShapeGradient {
    /// Construct iWork's two-stop simple linear gradient.
    pub fn linear(start: RgbaColor, end: RgbaColor, angle: ShapeGradientAngle) -> Self {
        Self {
            kind: ShapeGradientKind::Linear,
            stops: Box::new([
                ShapeGradientStop::new(
                    start,
                    ShapeGradientStopPosition::START,
                    ShapeGradientStopMidpoint::CENTER,
                ),
                ShapeGradientStop::new(
                    end,
                    ShapeGradientStopPosition::END,
                    ShapeGradientStopMidpoint::CENTER,
                ),
            ]),
            opacity: ShapeGradientOpacity::OPAQUE,
            is_advanced: false,
            angle,
        }
    }

    /// Construct an advanced linear or radial gradient.
    pub fn advanced(
        kind: ShapeGradientKind,
        stops: Vec<ShapeGradientStop>,
        opacity: ShapeGradientOpacity,
        angle: ShapeGradientAngle,
    ) -> Result<Self> {
        validate_stops(&stops)?;
        Ok(Self {
            kind,
            stops: stops.into_boxed_slice(),
            opacity,
            is_advanced: true,
            angle,
        })
    }

    pub const fn kind(&self) -> ShapeGradientKind {
        self.kind
    }

    pub fn stops(&self) -> &[ShapeGradientStop] {
        &self.stops
    }

    pub const fn opacity(&self) -> ShapeGradientOpacity {
        self.opacity
    }

    pub const fn is_advanced(&self) -> bool {
        self.is_advanced
    }

    pub const fn angle(&self) -> ShapeGradientAngle {
        self.angle
    }

    pub(crate) fn from_native_parts(
        kind: ShapeGradientKind,
        stops: Vec<ShapeGradientStop>,
        opacity: ShapeGradientOpacity,
        is_advanced: bool,
        angle: ShapeGradientAngle,
    ) -> Result<Self> {
        validate_stops(&stops)?;
        if !is_advanced
            && (kind != ShapeGradientKind::Linear
                || stops.len() != SIMPLE_STOP_COUNT
                || stops
                    .iter()
                    .any(|stop| stop.midpoint != ShapeGradientStopMidpoint::CENTER))
        {
            return Err(Error::InvalidFormat(
                "a simple iWork shape gradient must be linear with two centered stops".to_owned(),
            ));
        }
        Ok(Self {
            kind,
            stops: stops.into_boxed_slice(),
            opacity,
            is_advanced,
            angle,
        })
    }
}

fn validate_stops(stops: &[ShapeGradientStop]) -> Result<()> {
    if stops.len() < MINIMUM_STOP_COUNT {
        return Err(Error::ParseError(format!(
            "iWork shape gradients require at least {MINIMUM_STOP_COUNT} color stops"
        )));
    }
    if stops
        .windows(2)
        .any(|pair| pair[1].position < pair[0].position)
    {
        return Err(Error::ParseError(
            "iWork shape gradient stops must be ordered by position".to_owned(),
        ));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::shapes::RgbColorSpace;

    #[test]
    fn invalid_gradient_components_are_rejected() {
        assert!(ShapeGradientAngle::from_degrees(360.0).is_err());
        assert!(ShapeGradientAngle::from_degrees(f32::NAN).is_err());
        assert!(ShapeGradientStopPosition::new(-0.1).is_err());
        assert!(ShapeGradientStopMidpoint::new(1.1).is_err());
        assert!(ShapeGradientOpacity::new(f32::INFINITY).is_err());
    }

    #[test]
    fn advanced_gradient_requires_ordered_stops() {
        let color = RgbaColor::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap();
        let stops = vec![
            ShapeGradientStop::new(
                color,
                ShapeGradientStopPosition::END,
                ShapeGradientStopMidpoint::CENTER,
            ),
            ShapeGradientStop::new(
                color,
                ShapeGradientStopPosition::START,
                ShapeGradientStopMidpoint::CENTER,
            ),
        ];
        assert!(
            ShapeGradient::advanced(
                ShapeGradientKind::Linear,
                stops,
                ShapeGradientOpacity::OPAQUE,
                ShapeGradientAngle::from_degrees(0.0).unwrap(),
            )
            .is_err()
        );
    }
}
