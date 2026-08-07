//! Archive-independent RGBA color values used by iWork format owners.

/// The channel that failed [`Rgba::new`] validation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Channel {
    /// Red channel.
    Red,
    /// Green channel.
    Green,
    /// Blue channel.
    Blue,
    /// Alpha channel.
    Alpha,
}

impl std::fmt::Display for Channel {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::Red => "red",
            Self::Green => "green",
            Self::Blue => "blue",
            Self::Alpha => "alpha",
        })
    }
}

/// A channel value that is not a finite number in the inclusive `0.0..=1.0`
/// range.
#[derive(Debug, Clone, Copy, PartialEq, thiserror::Error)]
#[error("{channel} channel must be finite and between 0 and 1 (got {value})")]
pub struct Error {
    /// The invalid channel.
    pub channel: Channel,
    /// The rejected value.
    pub value: f32,
}

/// RGB color space used by native iWork colors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum RgbColorSpace {
    /// Standard RGB color space.
    #[default]
    Srgb,
    /// Display P3 wide-gamut color space.
    DisplayP3,
}

/// A validated normalized red, green, blue, and alpha color.
///
/// The value is copyable and contains no archive or protobuf state. Its fixed
/// representation keeps common style operations allocation-free.
#[repr(C)]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rgba {
    red: f32,
    green: f32,
    blue: f32,
    alpha: f32,
    color_space: RgbColorSpace,
}

impl Rgba {
    /// Construct a color whose channels are finite and in `0.0..=1.0`.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] identifying the first non-finite or out-of-range
    /// channel.
    pub fn new(
        red: f32,
        green: f32,
        blue: f32,
        alpha: f32,
        color_space: RgbColorSpace,
    ) -> Result<Self, Error> {
        for (channel, value) in [
            (Channel::Red, red),
            (Channel::Green, green),
            (Channel::Blue, blue),
            (Channel::Alpha, alpha),
        ] {
            if !value.is_finite() || !(0.0..=1.0).contains(&value) {
                return Err(Error { channel, value });
            }
        }
        Ok(Self {
            red,
            green,
            blue,
            alpha,
            color_space,
        })
    }

    /// Return the red channel.
    #[must_use]
    pub const fn red(self) -> f32 {
        self.red
    }

    /// Return the green channel.
    #[must_use]
    pub const fn green(self) -> f32 {
        self.green
    }

    /// Return the blue channel.
    #[must_use]
    pub const fn blue(self) -> f32 {
        self.blue
    }

    /// Return the alpha channel.
    #[must_use]
    pub const fn alpha(self) -> f32 {
        self.alpha
    }

    /// Return the color space.
    #[must_use]
    pub const fn color_space(self) -> RgbColorSpace {
        self.color_space
    }

    /// Return whether this color is fully opaque.
    #[must_use]
    #[allow(
        clippy::float_cmp,
        reason = "Opacity is the exact validated alpha endpoint, not an approximate measurement"
    )]
    pub const fn is_opaque(self) -> bool {
        self.alpha == 1.0
    }

    /// Return opaque sRGB black.
    #[must_use]
    pub const fn black() -> Self {
        Self {
            red: 0.0,
            green: 0.0,
            blue: 0.0,
            alpha: 1.0,
            color_space: RgbColorSpace::Srgb,
        }
    }

    /// Return transparent sRGB black, used by native text-outline defaults.
    #[must_use]
    pub const fn transparent_black() -> Self {
        Self {
            red: 0.0,
            green: 0.0,
            blue: 0.0,
            alpha: 0.0,
            color_space: RgbColorSpace::Srgb,
        }
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{Channel, RgbColorSpace, Rgba};

    #[test]
    fn value_is_compact_and_copyable() {
        assert_eq!(size_of::<Rgba>(), 20);
        let color = Rgba::black();
        assert_eq!(color, color);
        assert!(color.is_opaque());
        assert_eq!(
            Rgba::transparent_black().alpha().to_bits(),
            0.0_f32.to_bits()
        );
    }

    #[test]
    fn channels_are_validated_before_construction() {
        assert_eq!(
            Rgba::new(-0.1, 0.0, 0.0, 1.0, RgbColorSpace::Srgb)
                .map(|_| ())
                .err()
                .map(|error| error.channel),
            Some(Channel::Red)
        );
        assert_eq!(
            Rgba::new(0.0, f32::NAN, 0.0, 1.0, RgbColorSpace::Srgb)
                .map(|_| ())
                .err()
                .map(|error| error.channel),
            Some(Channel::Green)
        );
    }

    #[test]
    fn valid_values_preserve_channels_and_space() -> Result<(), super::Error> {
        let color = Rgba::new(0.1, 0.2, 0.3, 0.4, RgbColorSpace::DisplayP3)?;
        assert_eq!(color.red().to_bits(), 0.1_f32.to_bits());
        assert_eq!(color.green().to_bits(), 0.2_f32.to_bits());
        assert_eq!(color.blue().to_bits(), 0.3_f32.to_bits());
        assert_eq!(color.alpha().to_bits(), 0.4_f32.to_bits());
        assert_eq!(color.color_space(), RgbColorSpace::DisplayP3);
        assert!(!color.is_opaque());
        Ok(())
    }
}
