//! Archive-free semantic values for native iWork shape effects.

/// Validation failures for normalized shape-effect values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// Shape opacity was not finite.
    #[error("shape opacity must be finite")]
    OpacityNonFinite,
    /// Shape opacity was outside the inclusive normalized domain.
    #[error("shape opacity must be in 0.0..=1.0")]
    OpacityOutOfRange,
    /// Reflection opacity was not finite.
    #[error("shape reflection opacity must be finite")]
    ReflectionOpacityNonFinite,
    /// Reflection opacity was outside the inclusive normalized domain.
    #[error("shape reflection opacity must be in 0.0..=1.0")]
    ReflectionOpacityOutOfRange,
}

/// Result type for shape-effect value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// Validated normalized opacity of an entire shape, including its text and
/// stroke.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct Opacity(f32);

impl Opacity {
    /// Fully transparent shape opacity.
    pub const TRANSPARENT: Self = Self(0.0);
    /// Fully opaque shape opacity.
    pub const OPAQUE: Self = Self(1.0);

    /// Construct a finite opacity in the inclusive range `0.0..=1.0`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::OpacityNonFinite`] for NaN or infinity and
    /// [`Error::OpacityOutOfRange`] outside the inclusive normalized domain.
    #[must_use = "use the validated opacity or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(value, Error::OpacityNonFinite, Error::OpacityOutOfRange)?;
        Ok(Self(value))
    }

    /// Return the normalized opacity value.
    #[must_use]
    pub const fn value(self) -> f32 {
        self.0
    }
}

impl Default for Opacity {
    fn default() -> Self {
        Self::OPAQUE
    }
}

/// Validated normalized opacity of a reflected copy of a shape.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ReflectionOpacity(f32);

impl ReflectionOpacity {
    /// Fully invisible reflection opacity.
    pub const INVISIBLE: Self = Self(0.0);
    /// Native default reflection opacity.
    pub const DEFAULT: Self = Self(0.5);
    /// Fully opaque reflection opacity.
    pub const OPAQUE: Self = Self(1.0);

    /// Construct a finite reflection opacity in the inclusive range
    /// `0.0..=1.0`.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ReflectionOpacityNonFinite`] for NaN or infinity and
    /// [`Error::ReflectionOpacityOutOfRange`] outside the inclusive normalized
    /// domain.
    #[must_use = "use the validated reflection opacity or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        validate_normalized(
            value,
            Error::ReflectionOpacityNonFinite,
            Error::ReflectionOpacityOutOfRange,
        )?;
        Ok(Self(value))
    }

    /// Return the normalized reflection opacity value.
    #[must_use]
    pub const fn value(self) -> f32 {
        self.0
    }
}

impl Default for ReflectionOpacity {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Reflection state shown by the iWork Style inspector.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub enum Reflection {
    /// The reflection checkbox is off.
    #[default]
    Disabled,
    /// The reflection checkbox is on with a normalized opacity.
    Enabled(ReflectionOpacity),
}

/// Composable visual effects stored in an ordinary shape style.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq, Default)]
pub struct Effects {
    opacity: Opacity,
    reflection: Reflection,
}

impl Effects {
    /// Construct a complete shape-effects value.
    #[must_use]
    pub const fn new(opacity: Opacity, reflection: Reflection) -> Self {
        Self {
            opacity,
            reflection,
        }
    }

    /// Return the whole-shape opacity.
    #[must_use]
    pub const fn opacity(self) -> Opacity {
        self.opacity
    }

    /// Return the reflection state.
    #[must_use]
    pub const fn reflection(self) -> Reflection {
        self.reflection
    }

    /// Return effects with a different whole-shape opacity.
    #[must_use]
    pub const fn with_opacity(mut self, opacity: Opacity) -> Self {
        self.opacity = opacity;
        self
    }

    /// Return effects with a different reflection state.
    #[must_use]
    pub const fn with_reflection(mut self, reflection: Reflection) -> Self {
        self.reflection = reflection;
        self
    }
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
    use std::mem::{align_of, size_of};

    use super::{Effects, Error, Opacity, Reflection, ReflectionOpacity};

    #[test]
    fn effects_values_are_compact_and_strongly_typed() {
        assert_eq!(size_of::<Opacity>(), 4);
        assert_eq!(size_of::<ReflectionOpacity>(), 4);
        assert_eq!(size_of::<Reflection>(), 8);
        assert_eq!(size_of::<Effects>(), 12);
        assert_eq!(align_of::<Opacity>(), 4);
        assert_eq!(align_of::<ReflectionOpacity>(), 4);
        assert_eq!(align_of::<Effects>(), 4);

        let effects = Effects::default()
            .with_opacity(Opacity::TRANSPARENT)
            .with_reflection(Reflection::Enabled(ReflectionOpacity::DEFAULT));
        assert_eq!(effects.opacity(), Opacity::TRANSPARENT);
        assert_eq!(
            effects.reflection(),
            Reflection::Enabled(ReflectionOpacity::DEFAULT)
        );
    }

    #[test]
    fn leaf_validation_reports_typed_failures() {
        assert_eq!(Opacity::new(f32::NAN), Err(Error::OpacityNonFinite));
        assert_eq!(Opacity::new(f32::INFINITY), Err(Error::OpacityNonFinite));
        assert_eq!(Opacity::new(-0.01), Err(Error::OpacityOutOfRange));
        assert_eq!(Opacity::new(1.01), Err(Error::OpacityOutOfRange));
        assert_eq!(
            ReflectionOpacity::new(f32::NEG_INFINITY),
            Err(Error::ReflectionOpacityNonFinite)
        );
        assert_eq!(
            ReflectionOpacity::new(-0.01),
            Err(Error::ReflectionOpacityOutOfRange)
        );
        assert_eq!(
            ReflectionOpacity::new(1.01),
            Err(Error::ReflectionOpacityOutOfRange)
        );
    }

    #[test]
    fn valid_leaf_values_preserve_endpoints_and_negative_zero() {
        assert_eq!(
            Opacity::new(-0.0)
                .unwrap_or_else(|_| panic!("negative zero is a valid opacity"))
                .value()
                .to_bits(),
            (-0.0f32).to_bits()
        );
        assert_eq!(Opacity::OPAQUE.value().to_bits(), 1.0f32.to_bits());
        assert_eq!(
            ReflectionOpacity::new(0.35)
                .unwrap_or_else(|_| panic!("valid reflection opacity"))
                .value()
                .to_bits(),
            0.35f32.to_bits()
        );
    }
}
