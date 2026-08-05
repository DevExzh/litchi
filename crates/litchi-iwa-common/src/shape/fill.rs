//! Archive-free shape-gradient values shared by iWork format owners.

use std::f32::consts::TAU;

use crate::color::Rgba;

const MINIMUM_STOP_COUNT: usize = 2;
const SIMPLE_STOP_COUNT: usize = 2;

/// Validation failures for shape-gradient values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The gradient angle was not finite.
    #[error("shape gradient angle must be finite")]
    AngleNonFinite,
    /// The gradient angle was outside the native `[0, 2π)` interval.
    #[error("shape gradient angle must be in [0, 2π)")]
    AngleOutOfRange,
    /// A stop position was not finite.
    #[error("shape gradient stop position must be finite")]
    StopPositionNonFinite,
    /// A stop position was outside the normalized interval.
    #[error("shape gradient stop position must be in 0.0..=1.0")]
    StopPositionOutOfRange,
    /// A stop midpoint was not finite.
    #[error("shape gradient stop midpoint must be finite")]
    StopMidpointNonFinite,
    /// A stop midpoint was outside the normalized interval.
    #[error("shape gradient stop midpoint must be in 0.0..=1.0")]
    StopMidpointOutOfRange,
    /// Gradient opacity was not finite.
    #[error("shape gradient opacity must be finite")]
    OpacityNonFinite,
    /// Gradient opacity was outside the normalized interval.
    #[error("shape gradient opacity must be in 0.0..=1.0")]
    OpacityOutOfRange,
    /// A gradient contained fewer than two stops.
    #[error("shape gradients require at least {MINIMUM_STOP_COUNT} color stops")]
    TooFewStops,
    /// Stops were not ordered by their normalized position.
    #[error("shape gradient stops must be ordered by position")]
    StopsUnordered,
    /// A non-advanced gradient did not have the native simple shape.
    #[error("a simple shape gradient must be linear with two centered stops")]
    InvalidSimpleShape,
}

/// Result type for shape-gradient value construction.
pub type Result<T> = std::result::Result<T, Error>;

#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The macro keeps three identical validated scalar declarations concise"
)]
macro_rules! normalized_value {
    ($name:ident, $non_finite:ident, $out_of_range:ident, $doc:literal) => {
        #[doc = $doc]
        #[repr(transparent)]
        #[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
        pub struct $name(f32);

        impl $name {
            /// Construct a finite normalized value in the inclusive range `0.0..=1.0`.
            ///
            /// # Errors
            ///
            /// Returns a typed error when `value` is non-finite or outside
            /// the normalized interval.
            pub fn new(value: f32) -> Result<Self> {
                if !value.is_finite() {
                    return Err(Error::$non_finite);
                }
                if !(0.0..=1.0).contains(&value) {
                    return Err(Error::$out_of_range);
                }
                Ok(Self(value))
            }

            /// Return the normalized value.
            #[must_use]
            pub const fn get(self) -> f32 {
                self.0
            }
        }
    };
}

/// Geometry used to paint a shape gradient.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// Colors progress along a straight axis.
    Linear,
    /// Colors radiate from the native gradient origin.
    Radial,
}

/// Angle of a shape gradient, stored in native radians.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Angle(f32);

impl Angle {
    /// Construct an angle from the degree value shown by iWork.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the converted angle is non-finite or
    /// outside the native `[0, 2π)` interval.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        Self::from_radians(degrees.to_radians())
    }

    /// Construct an angle from the native radian value.
    ///
    /// # Errors
    ///
    /// Returns a typed error when `radians` is non-finite or outside the
    /// native `[0, 2π)` interval.
    pub fn from_radians(radians: f32) -> Result<Self> {
        if !radians.is_finite() {
            return Err(Error::AngleNonFinite);
        }
        if !(0.0..TAU).contains(&radians) {
            return Err(Error::AngleOutOfRange);
        }
        Ok(Self(radians))
    }

    /// Return the exact radian value stored in the native archive.
    #[must_use]
    pub const fn radians(self) -> f32 {
        self.0
    }

    /// Return the angle in the degree unit displayed by iWork.
    #[must_use]
    pub fn degrees(self) -> f32 {
        self.0.to_degrees()
    }
}

normalized_value!(
    StopPosition,
    StopPositionNonFinite,
    StopPositionOutOfRange,
    "Normalized location of a shape-gradient color stop."
);
normalized_value!(
    StopMidpoint,
    StopMidpointNonFinite,
    StopMidpointOutOfRange,
    "Normalized blend midpoint following a shape-gradient color stop."
);
normalized_value!(
    Opacity,
    OpacityNonFinite,
    OpacityOutOfRange,
    "Normalized overall opacity of a shape gradient."
);

impl StopPosition {
    /// The beginning of a gradient axis.
    pub const START: Self = Self(0.0);
    /// The end of a gradient axis.
    pub const END: Self = Self(1.0);
}

impl StopMidpoint {
    /// The midpoint used by iWork's simple two-color gradient editor.
    pub const CENTER: Self = Self(0.5);
}

impl Opacity {
    /// A fully opaque gradient.
    pub const OPAQUE: Self = Self(1.0);
}

/// One immutable, validated color stop in a shape gradient.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Stop {
    color: Rgba,
    position: StopPosition,
    midpoint: StopMidpoint,
}

impl Stop {
    /// Construct one color stop from validated components.
    #[must_use]
    pub const fn new(color: Rgba, position: StopPosition, midpoint: StopMidpoint) -> Self {
        Self {
            color,
            position,
            midpoint,
        }
    }

    /// Return the stop color.
    #[must_use]
    pub const fn color(self) -> Rgba {
        self.color
    }

    /// Return the normalized stop position.
    #[must_use]
    pub const fn position(self) -> StopPosition {
        self.position
    }

    /// Return the normalized stop midpoint.
    #[must_use]
    pub const fn midpoint(self) -> StopMidpoint {
        self.midpoint
    }
}

/// Native simple or advanced linear/radial shape gradient.
#[derive(Debug, Clone, PartialEq)]
pub struct Gradient {
    kind: Kind,
    stops: Box<[Stop]>,
    opacity: Opacity,
    advanced: bool,
    angle: Angle,
}

impl Gradient {
    /// Construct iWork's two-stop simple linear gradient.
    #[must_use]
    pub fn linear(start: Rgba, end: Rgba, angle: Angle) -> Self {
        Self {
            kind: Kind::Linear,
            stops: Box::new([
                Stop::new(start, StopPosition::START, StopMidpoint::CENTER),
                Stop::new(end, StopPosition::END, StopMidpoint::CENTER),
            ]),
            opacity: Opacity::OPAQUE,
            advanced: false,
            angle,
        }
    }

    /// Construct an explicitly advanced linear or radial gradient.
    ///
    /// # Errors
    ///
    /// Returns a typed error when fewer than two stops are supplied or their
    /// positions are not ordered.
    pub fn advanced(
        kind: Kind,
        stops: impl Into<Box<[Stop]>>,
        opacity: Opacity,
        angle: Angle,
    ) -> Result<Self> {
        Self::from_parts(kind, stops, opacity, true, angle)
    }

    /// Construct a gradient from all native semantic fields.
    ///
    /// The `advanced` flag is retained because it changes how iWork presents
    /// and edits the same gradient wire shape; it is not inferred from the
    /// stop count so a read/write cycle remains lossless.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the stops are insufficient or unordered, or
    /// when a simple gradient does not have its required native shape.
    pub fn from_parts(
        kind: Kind,
        stops: impl Into<Box<[Stop]>>,
        opacity: Opacity,
        advanced: bool,
        angle: Angle,
    ) -> Result<Self> {
        let boxed_stops = stops.into();
        validate_stops(&boxed_stops)?;
        if !advanced
            && (kind != Kind::Linear
                || boxed_stops.len() != SIMPLE_STOP_COUNT
                || boxed_stops
                    .iter()
                    .any(|stop| stop.midpoint() != StopMidpoint::CENTER))
        {
            return Err(Error::InvalidSimpleShape);
        }
        Ok(Self {
            kind,
            stops: boxed_stops,
            opacity,
            advanced,
            angle,
        })
    }

    /// Return the gradient geometry.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the boxed stop slice without copying it.
    #[must_use]
    pub fn stops(&self) -> &[Stop] {
        &self.stops
    }

    /// Return the overall gradient opacity.
    #[must_use]
    pub const fn opacity(&self) -> Opacity {
        self.opacity
    }

    /// Return whether iWork stores this as an advanced gradient.
    #[must_use]
    pub const fn is_advanced(&self) -> bool {
        self.advanced
    }

    /// Return the gradient angle.
    #[must_use]
    pub const fn angle(&self) -> Angle {
        self.angle
    }
}

fn validate_stops(stops: &[Stop]) -> Result<()> {
    if stops.len() < MINIMUM_STOP_COUNT {
        return Err(Error::TooFewStops);
    }
    if stops
        .windows(2)
        .any(|pair| pair[1].position() < pair[0].position())
    {
        return Err(Error::StopsUnordered);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::*;
    use crate::color::RgbColorSpace;

    fn color() -> Rgba {
        Rgba::new(0.2, 0.4, 0.8, 1.0, RgbColorSpace::Srgb).unwrap()
    }

    #[test]
    fn scalar_values_are_typed_and_compact() {
        assert_eq!(size_of::<Angle>(), 4);
        assert_eq!(size_of::<StopPosition>(), 4);
        assert_eq!(size_of::<StopMidpoint>(), 4);
        assert_eq!(size_of::<Opacity>(), 4);
        assert_eq!(align_of::<Angle>(), 4);
        assert_eq!(size_of::<Stop>(), size_of::<Rgba>() + 8);
    }

    #[test]
    fn scalar_validation_is_typed() {
        assert_eq!(Angle::from_degrees(f32::NAN), Err(Error::AngleNonFinite));
        assert_eq!(Angle::from_degrees(360.0), Err(Error::AngleOutOfRange));
        assert_eq!(StopPosition::new(-0.1), Err(Error::StopPositionOutOfRange));
        assert_eq!(
            StopMidpoint::new(f32::INFINITY),
            Err(Error::StopMidpointNonFinite)
        );
        assert_eq!(Opacity::new(1.1), Err(Error::OpacityOutOfRange));
    }

    #[test]
    fn gradients_keep_boxed_stops_and_native_mode() {
        let angle = Angle::from_degrees(45.0).unwrap();
        let gradient = Gradient::from_parts(
            Kind::Radial,
            vec![
                Stop::new(color(), StopPosition::START, StopMidpoint::CENTER),
                Stop::new(color(), StopPosition::END, StopMidpoint::CENTER),
            ],
            Opacity::new(0.8).unwrap(),
            true,
            angle,
        )
        .unwrap();
        assert!(gradient.is_advanced());
        assert_eq!(gradient.stops().len(), 2);
        assert_eq!(gradient.angle(), angle);
    }

    #[test]
    fn gradients_reject_invalid_stop_sequences() {
        let color = color();
        let stops = vec![
            Stop::new(color, StopPosition::END, StopMidpoint::CENTER),
            Stop::new(color, StopPosition::START, StopMidpoint::CENTER),
        ];
        assert_eq!(
            Gradient::advanced(
                Kind::Linear,
                stops,
                Opacity::OPAQUE,
                Angle::from_degrees(0.0).unwrap(),
            ),
            Err(Error::StopsUnordered)
        );
    }
}
