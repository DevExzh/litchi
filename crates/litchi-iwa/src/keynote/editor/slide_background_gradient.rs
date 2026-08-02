//! Strict public model and validation for Keynote gradient backgrounds.

use std::f32::consts::TAU;

use super::slide_background_color::{KeynoteRgbaColor, validate_color};
use super::*;

const MINIMUM_STOP_COUNT: usize = 2;
const SIMPLE_STOP_COUNT: usize = 2;
const DEFAULT_MIDPOINT: f32 = 0.5;
const FULL_OPACITY: f32 = 1.0;

/// Native Keynote gradient geometry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteGradientKind {
    /// Colors progress along a straight axis.
    Linear,
    /// Colors radiate from the native gradient origin.
    Radial,
}

/// Angle of a Keynote gradient, stored in native radians.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteGradientAngle(f32);

impl KeynoteGradientAngle {
    /// Construct an angle from the degree value shown by Keynote.
    pub fn from_degrees(degrees: f32) -> Result<Self> {
        Self::from_radians(degrees.to_radians())
    }

    /// Construct an angle from Keynote's native radian value.
    pub fn from_radians(radians: f32) -> Result<Self> {
        if !radians.is_finite() || !(0.0..TAU).contains(&radians) {
            return Err(Error::ParseError(
                "Keynote gradient angle must be finite and in [0, 2π) radians".to_owned(),
            ));
        }
        Ok(Self(radians))
    }

    /// Return the exact radian value stored in the native archive.
    pub fn radians(self) -> f32 {
        self.0
    }

    /// Return the angle in the degree unit displayed by Keynote.
    pub fn degrees(self) -> f32 {
        self.0.to_degrees()
    }
}

/// One ordered color stop in a Keynote gradient.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteGradientStop {
    pub color: KeynoteRgbaColor,
    /// Normalized location in the inclusive range `0..=1`.
    pub position: f32,
    /// Normalized blend midpoint in the inclusive range `0..=1`.
    pub midpoint: f32,
}

impl KeynoteGradientStop {
    /// Construct and validate one normalized gradient stop.
    pub fn new(color: KeynoteRgbaColor, position: f32, midpoint: f32) -> Result<Self> {
        let stop = Self {
            color,
            position,
            midpoint,
        };
        validate_stop(&stop, 0)?;
        Ok(stop)
    }
}

/// Typed native Keynote gradient fill.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteGradient {
    /// Linear or radial native geometry.
    pub kind: KeynoteGradientKind,
    /// Ordered color stops; at least two are required.
    pub stops: Vec<KeynoteGradientStop>,
    /// Normalized opacity in the inclusive range `0..=1`.
    pub opacity: f32,
    /// Whether Keynote should expose advanced stop controls.
    pub is_advanced: bool,
    /// Direction shown in Keynote's angle control.
    pub angle: KeynoteGradientAngle,
}

impl KeynoteGradient {
    /// Construct Keynote's two-stop simple linear gradient.
    pub fn linear(
        start: KeynoteRgbaColor,
        end: KeynoteRgbaColor,
        angle: KeynoteGradientAngle,
    ) -> Result<Self> {
        let gradient = Self {
            kind: KeynoteGradientKind::Linear,
            stops: vec![
                KeynoteGradientStop {
                    color: start,
                    position: 0.0,
                    midpoint: DEFAULT_MIDPOINT,
                },
                KeynoteGradientStop {
                    color: end,
                    position: 1.0,
                    midpoint: DEFAULT_MIDPOINT,
                },
            ],
            opacity: FULL_OPACITY,
            is_advanced: false,
            angle,
        };
        validate_gradient(&gradient)?;
        Ok(gradient)
    }

    /// Construct and validate an advanced linear or radial gradient.
    pub fn advanced(
        kind: KeynoteGradientKind,
        stops: Vec<KeynoteGradientStop>,
        opacity: f32,
        angle: KeynoteGradientAngle,
    ) -> Result<Self> {
        let gradient = Self {
            kind,
            stops,
            opacity,
            is_advanced: true,
            angle,
        };
        validate_gradient(&gradient)?;
        Ok(gradient)
    }
}

pub(super) fn validate_gradient(gradient: &KeynoteGradient) -> Result<()> {
    if gradient.stops.len() < MINIMUM_STOP_COUNT {
        return Err(Error::ParseError(format!(
            "Keynote gradients require at least {MINIMUM_STOP_COUNT} color stops"
        )));
    }
    if !gradient.is_advanced
        && (gradient.kind != KeynoteGradientKind::Linear
            || gradient.stops.len() != SIMPLE_STOP_COUNT
            || gradient
                .stops
                .iter()
                .any(|stop| stop.midpoint != DEFAULT_MIDPOINT))
    {
        return Err(Error::ParseError(
            "a simple Keynote gradient must be linear with two 0.5-midpoint stops".to_owned(),
        ));
    }
    if !gradient.opacity.is_finite() || !(0.0..=1.0).contains(&gradient.opacity) {
        return Err(Error::ParseError(
            "Keynote gradient opacity must be finite and between 0 and 1".to_owned(),
        ));
    }
    let mut previous_position = None;
    for (index, stop) in gradient.stops.iter().enumerate() {
        validate_stop(stop, index)?;
        if previous_position.is_some_and(|previous| stop.position < previous) {
            return Err(Error::ParseError(
                "Keynote gradient stops must be ordered by position".to_owned(),
            ));
        }
        previous_position = Some(stop.position);
    }
    Ok(())
}

fn validate_stop(stop: &KeynoteGradientStop, index: usize) -> Result<()> {
    validate_color(stop.color, "Keynote gradient-stop")?;
    for (name, value) in [("position", stop.position), ("midpoint", stop.midpoint)] {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(Error::ParseError(format!(
                "Keynote gradient stop {index} {name} must be finite and between 0 and 1"
            )));
        }
    }
    Ok(())
}
