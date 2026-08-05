//! Archive-free image-adjustment values shared by iWork format owners.
//!
//! The native image archive and its wire representation stay in the concrete
//! `litchi-iwa` adapter. This module owns only the validated semantic values
//! exchanged at that boundary. In particular, an omitted control is distinct
//! from a control explicitly set to its neutral value.

#![allow(
    clippy::module_name_repetitions,
    reason = "The public Image* names identify the native inspector controls while staying concise at the shape::image boundary"
)]

const EXPOSURE_PRESENT: u8 = 1 << 0;
const SATURATION_PRESENT: u8 = 1 << 1;
const ENHANCEMENT_PRESENT: u8 = 1 << 2;

/// Validation failures for normalized image-adjustment values.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// The supplied adjustment was NaN or infinite.
    #[error("image adjustment must be finite")]
    AdjustmentNonFinite,
    /// The supplied adjustment was outside the inclusive native domain.
    #[error("image adjustment must be in -1.0..=1.0")]
    AdjustmentOutOfRange,
}

/// Result type for image-adjustment value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A normalized native image-adjustment amount.
///
/// iWork exposes exposure and saturation as percentages. Values are stored as
/// a normalized multiplier in the inclusive range `-1.0..=1.0`, where `0.25`
/// corresponds to `25%` in the native inspector.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, PartialOrd)]
pub struct ImageAdjustment(f32);

impl ImageAdjustment {
    /// The lowest native adjustment amount (`-100%`).
    pub const MINIMUM: Self = Self(-1.0);
    /// The neutral adjustment amount (`0%`).
    pub const NEUTRAL: Self = Self(0.0);
    /// The highest native adjustment amount (`100%`).
    pub const MAXIMUM: Self = Self(1.0);

    /// Construct one finite native image-adjustment amount.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonFinite`] for NaN or infinity and
    /// [`Error::OutOfRange`] outside the inclusive `-1.0..=1.0` domain.
    #[must_use = "use the validated adjustment or handle its validation error"]
    pub fn new(value: f32) -> Result<Self> {
        if !value.is_finite() {
            return Err(Error::AdjustmentNonFinite);
        }
        if !(-1.0..=1.0).contains(&value) {
            return Err(Error::AdjustmentOutOfRange);
        }
        Ok(Self(value))
    }

    /// Return the normalized native amount.
    #[must_use]
    pub const fn value(self) -> f32 {
        self.0
    }
}

impl TryFrom<f32> for ImageAdjustment {
    type Error = Error;

    fn try_from(value: f32) -> Result<Self> {
        Self::new(value)
    }
}

/// State of iWork's native Enhance control.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum ImageEnhancement {
    /// Leave automatic enhancement disabled.
    #[default]
    Disabled,
    /// Let iWork automatically enhance the image colors.
    Enabled,
}

/// The basic controls in iWork's Image inspector.
///
/// Each optional field preserves native presence. For example,
/// `with_exposure(None)` means the exposure field is omitted, while
/// `with_exposure(Some(ImageAdjustment::NEUTRAL))` means it is explicitly
/// encoded as neutral. Advanced native adjustment fields are intentionally
/// outside this focused semantic value.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageAdjustments {
    exposure: f32,
    saturation: f32,
    enhancement: ImageEnhancement,
    present: u8,
}

impl ImageAdjustments {
    /// Construct adjustments with every basic control omitted.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            exposure: 0.0,
            saturation: 0.0,
            enhancement: ImageEnhancement::Disabled,
            present: 0,
        }
    }

    /// Return the optional exposure adjustment.
    #[must_use]
    pub const fn exposure(self) -> Option<ImageAdjustment> {
        if self.present & EXPOSURE_PRESENT != 0 {
            Some(ImageAdjustment(self.exposure))
        } else {
            None
        }
    }

    /// Return the optional saturation adjustment.
    #[must_use]
    pub const fn saturation(self) -> Option<ImageAdjustment> {
        if self.present & SATURATION_PRESENT != 0 {
            Some(ImageAdjustment(self.saturation))
        } else {
            None
        }
    }

    /// Return the optional automatic-enhancement state.
    #[must_use]
    pub const fn enhancement(self) -> Option<ImageEnhancement> {
        if self.present & ENHANCEMENT_PRESENT != 0 {
            Some(self.enhancement)
        } else {
            None
        }
    }

    /// Return adjustments with the optional exposure control replaced.
    #[must_use]
    pub const fn with_exposure(mut self, exposure: Option<ImageAdjustment>) -> Self {
        if let Some(adjustment) = exposure {
            self.exposure = adjustment.value();
            self.present |= EXPOSURE_PRESENT;
        } else {
            self.exposure = 0.0;
            self.present &= !EXPOSURE_PRESENT;
        }
        self
    }

    /// Return adjustments with the optional saturation control replaced.
    #[must_use]
    pub const fn with_saturation(mut self, saturation: Option<ImageAdjustment>) -> Self {
        if let Some(adjustment) = saturation {
            self.saturation = adjustment.value();
            self.present |= SATURATION_PRESENT;
        } else {
            self.saturation = 0.0;
            self.present &= !SATURATION_PRESENT;
        }
        self
    }

    /// Return adjustments with the optional automatic-enhancement control
    /// replaced.
    #[must_use]
    pub const fn with_enhancement(mut self, enhancement: Option<ImageEnhancement>) -> Self {
        if let Some(state) = enhancement {
            self.enhancement = state;
            self.present |= ENHANCEMENT_PRESENT;
        } else {
            self.enhancement = ImageEnhancement::Disabled;
            self.present &= !ENHANCEMENT_PRESENT;
        }
        self
    }
}

impl Default for ImageAdjustments {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::{align_of, size_of};

    use super::{Error, ImageAdjustment, ImageAdjustments, ImageEnhancement};

    #[test]
    fn adjustment_is_compact_and_strictly_validated() {
        assert_eq!(size_of::<ImageAdjustment>(), 4);
        assert_eq!(align_of::<ImageAdjustment>(), 4);
        assert_eq!(
            ImageAdjustment::new(f32::NAN),
            Err(Error::AdjustmentNonFinite)
        );
        assert_eq!(
            ImageAdjustment::new(f32::NEG_INFINITY),
            Err(Error::AdjustmentNonFinite)
        );
        assert_eq!(
            ImageAdjustment::new(f32::INFINITY),
            Err(Error::AdjustmentNonFinite)
        );
        assert_eq!(
            ImageAdjustment::new(-1.01),
            Err(Error::AdjustmentOutOfRange)
        );
        assert_eq!(ImageAdjustment::new(1.01), Err(Error::AdjustmentOutOfRange));
        assert_eq!(
            ImageAdjustment::new(-1.0).unwrap(),
            ImageAdjustment::MINIMUM
        );
        assert_eq!(ImageAdjustment::new(1.0).unwrap(), ImageAdjustment::MAXIMUM);
        assert_eq!(ImageAdjustment::try_from(0.25).unwrap().value(), 0.25);
    }

    #[test]
    fn adjustments_are_copyable_and_keep_presence_distinct_from_neutral() {
        assert_eq!(size_of::<ImageEnhancement>(), 1);
        assert_eq!(size_of::<ImageAdjustments>(), 12);

        let omitted = ImageAdjustments::new();
        assert_eq!(omitted, ImageAdjustments::default());
        assert_eq!(omitted.exposure(), None);
        assert_eq!(omitted.saturation(), None);
        assert_eq!(omitted.enhancement(), None);

        let explicit_neutral = omitted.with_exposure(Some(ImageAdjustment::NEUTRAL));
        assert_eq!(explicit_neutral.exposure(), Some(ImageAdjustment::NEUTRAL));
        assert_ne!(explicit_neutral, omitted);
        assert_eq!(omitted.exposure(), None);

        let explicit_saturation = omitted.with_saturation(Some(ImageAdjustment::NEUTRAL));
        assert_eq!(
            explicit_saturation.saturation(),
            Some(ImageAdjustment::NEUTRAL)
        );
        assert_eq!(omitted.saturation(), None);

        let explicit_disabled = omitted.with_enhancement(Some(ImageEnhancement::Disabled));
        assert_eq!(
            explicit_disabled.enhancement(),
            Some(ImageEnhancement::Disabled)
        );
        assert_eq!(omitted.enhancement(), None);
    }

    #[test]
    fn immutable_builders_replace_only_the_requested_control() {
        let base = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let cleared = base.with_exposure(None);

        assert_eq!(base.exposure(), Some(ImageAdjustment::new(0.25).unwrap()));
        assert_eq!(cleared.exposure(), None);
        assert_eq!(cleared.saturation(), base.saturation());
        assert_eq!(cleared.enhancement(), base.enhancement());
        assert_eq!(ImageEnhancement::default(), ImageEnhancement::Disabled);
    }
}
