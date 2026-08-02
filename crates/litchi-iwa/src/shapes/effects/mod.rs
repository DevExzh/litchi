//! Typed object opacity and reflection effects for ordinary shapes.

mod native;
mod style;

use crate::{Error, Result};

pub(crate) use style::{reset_shape_effects, set_shape_effects, shape_effects};

/// Normalized opacity of an entire shape, including its text and stroke.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeOpacity(f32);

impl ShapeOpacity {
    pub const TRANSPARENT: Self = Self(0.0);
    pub const OPAQUE: Self = Self(1.0);

    /// Construct a finite opacity in the inclusive range `0.0..=1.0`.
    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(value, "Shape opacity")?;
        Ok(Self(value))
    }

    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for ShapeOpacity {
    fn default() -> Self {
        Self::OPAQUE
    }
}

/// Normalized opacity of a reflected copy of a shape.
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ShapeReflectionOpacity(f32);

impl ShapeReflectionOpacity {
    pub const INVISIBLE: Self = Self(0.0);
    pub const DEFAULT: Self = Self(0.5);
    pub const OPAQUE: Self = Self(1.0);

    /// Construct a finite reflection opacity in the inclusive range `0.0..=1.0`.
    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(value, "Shape reflection opacity")?;
        Ok(Self(value))
    }

    pub const fn get(self) -> f32 {
        self.0
    }
}

impl Default for ShapeReflectionOpacity {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Reflection state shown by the iWork Style inspector.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum ShapeReflection {
    /// The reflection checkbox is off.
    #[default]
    Disabled,
    /// The reflection checkbox is on with a normalized opacity.
    Enabled(ShapeReflectionOpacity),
}

/// Composable visual effects stored in an ordinary shape style.
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct ShapeEffects {
    opacity: ShapeOpacity,
    reflection: ShapeReflection,
}

impl ShapeEffects {
    pub const fn new(opacity: ShapeOpacity, reflection: ShapeReflection) -> Self {
        Self {
            opacity,
            reflection,
        }
    }

    pub const fn opacity(self) -> ShapeOpacity {
        self.opacity
    }

    pub const fn reflection(self) -> ShapeReflection {
        self.reflection
    }

    pub const fn with_opacity(mut self, opacity: ShapeOpacity) -> Self {
        self.opacity = opacity;
        self
    }

    pub const fn with_reflection(mut self, reflection: ShapeReflection) -> Self {
        self.reflection = reflection;
        self
    }
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
    fn normalized_effect_values_reject_invalid_inputs() {
        for invalid in [f32::NAN, f32::INFINITY, -0.01, 1.01] {
            assert!(ShapeOpacity::new(invalid).is_err());
            assert!(ShapeReflectionOpacity::new(invalid).is_err());
        }
    }

    #[test]
    fn effect_builders_preserve_strong_types() {
        let opacity = ShapeOpacity::new(0.72).unwrap();
        let reflection = ShapeReflectionOpacity::new(0.35).unwrap();
        let effects = ShapeEffects::default()
            .with_opacity(opacity)
            .with_reflection(ShapeReflection::Enabled(reflection));
        assert_eq!(effects.opacity(), opacity);
        assert_eq!(effects.reflection(), ShapeReflection::Enabled(reflection));
    }
}
