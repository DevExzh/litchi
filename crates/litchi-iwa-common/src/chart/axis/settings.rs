//! Archive-free settings for a native chart value axis.
//!
//! This aggregate deliberately composes the individually validated
//! [`Bounds`], [`Steps`], and [`Scale`] values. It owns no package, archive,
//! protobuf, or native-identifier state; concrete format adapters are
//! responsible for translating the value to and from their private native
//! representation and for enforcing any format-specific cross-setting rules.

use super::{Bounds, Scale, Steps};

/// The complete semantic settings of a chart's primary value axis.
///
/// The three members are strict semantic values rather than raw native
/// scalars. [`Bounds`] rejects non-finite and inverted manual ranges, and
/// [`Steps`] rejects zero major counts and values outside the native signed
/// integer range before they can enter this aggregate. [`Scale`] intentionally
/// accepts [`Scale::Unsupported`] so a reader can preserve a value introduced
/// by a newer iWork release. Consequently, [`Self::new`] is infallible with
/// respect to component validity. It does not assert format-specific
/// cross-setting rules—for example, whether a particular native format
/// permits a logarithmic scale with a given range—so concrete adapters must
/// validate those combinations before publication.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ValueAxisSettings {
    bounds: Bounds,
    steps: Steps,
    scale: Scale,
}

impl ValueAxisSettings {
    /// Build complete value-axis settings from validated semantic values.
    ///
    /// The constructor is intentionally infallible. Callers obtain bounds
    /// through [`Bounds::new`], [`Bounds::automatic`], or [`Bounds::fixed`]
    /// and step counts through their checked constructors. An unrecognized
    /// native scale remains valid as [`Scale::Unsupported`] and is preserved
    /// by a format adapter rather than rejected or mapped to a known default.
    #[must_use]
    pub const fn new(bounds: Bounds, steps: Steps, scale: Scale) -> Self {
        Self {
            bounds,
            steps,
            scale,
        }
    }

    /// Use automatic bounds and step counts with the native linear scale.
    #[must_use]
    pub const fn automatic() -> Self {
        Self::new(Bounds::automatic(), Steps::automatic(), Scale::Linear)
    }

    /// Return the manual or automatic value-axis bounds.
    #[must_use]
    pub const fn bounds(self) -> Bounds {
        self.bounds
    }

    /// Return the manual or automatic value-axis step counts.
    #[must_use]
    pub const fn steps(self) -> Steps {
        self.steps
    }

    /// Return the value-axis numeric scale.
    #[must_use]
    pub const fn scale(self) -> Scale {
        self.scale
    }

    /// Return settings with replacement bounds.
    #[must_use]
    pub const fn with_bounds(mut self, bounds: Bounds) -> Self {
        self.bounds = bounds;
        self
    }

    /// Return settings with replacement step counts.
    #[must_use]
    pub const fn with_steps(mut self, steps: Steps) -> Self {
        self.steps = steps;
        self
    }

    /// Return settings with a replacement numeric scale.
    #[must_use]
    pub const fn with_scale(mut self, scale: Scale) -> Self {
        self.scale = scale;
        self
    }
}

impl Default for ValueAxisSettings {
    fn default() -> Self {
        Self::automatic()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::ValueAxisSettings;
    use crate::chart::axis::{Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps};

    #[test]
    fn settings_are_compact_and_archive_free() {
        assert_eq!(size_of::<ValueAxisSettings>(), 48);
        assert_eq!(align_of::<ValueAxisSettings>(), 8);
        assert_eq!(ValueAxisSettings::default(), ValueAxisSettings::automatic());
        assert_eq!(
            ValueAxisSettings::automatic(),
            ValueAxisSettings::new(Bounds::automatic(), Steps::automatic(), Scale::Linear)
        );
    }

    #[test]
    fn settings_getters_and_consuming_builders_preserve_components() {
        let minimum = Bound::new(-10.0).unwrap();
        let maximum = Bound::new(100.0).unwrap();
        let bounds = Bounds::fixed(minimum, maximum).unwrap();
        let steps = Steps::fixed(
            MajorStepCount::new(5).unwrap(),
            MinorStepCount::new(2).unwrap(),
        );
        let settings = ValueAxisSettings::default()
            .with_bounds(bounds)
            .with_steps(steps)
            .with_scale(Scale::Logarithmic);

        assert_eq!(settings.bounds(), bounds);
        assert_eq!(settings.steps(), steps);
        assert_eq!(settings.scale(), Scale::Logarithmic);
    }

    #[test]
    fn unsupported_scale_values_are_valid_and_lossless() {
        let settings = ValueAxisSettings::automatic().with_scale(Scale::Unsupported(9_001));

        assert_eq!(settings.scale(), Scale::Unsupported(9_001));
        assert_eq!(settings.scale().native_value(), 9_001);
    }

    #[test]
    fn construction_accepts_only_validated_component_values() {
        let bounds = Bounds::new(None, Bound::new(10.0).ok()).unwrap();
        let steps = Steps::new(MajorStepCount::new(1).ok(), MinorStepCount::new(0).ok());
        let settings = ValueAxisSettings::new(bounds, steps, Scale::Linear);

        assert_eq!(settings.bounds().maximum().unwrap().value(), 10.0);
        assert_eq!(settings.steps().major().unwrap().value(), 1);
        assert_eq!(settings.steps().minor().unwrap().value(), 0);
    }

    #[test]
    fn bounds_presence_preserves_explicit_zero_and_partial_ranges() {
        let explicit_zero = Bound::new(0.0).unwrap();
        let partial = Bounds::new(Some(explicit_zero), None).unwrap();
        let settings = ValueAxisSettings::automatic().with_bounds(partial);

        assert_ne!(partial, Bounds::automatic());
        assert_eq!(partial.minimum(), Some(explicit_zero));
        assert_eq!(partial.maximum(), None);
        assert_eq!(settings.bounds(), partial);
    }
}
