//! Validated drawing colors and native RGB conversion.

use crate::protobuf::tsp;
use crate::{Error, Result};

/// RGB color space used by native iWork drawing colors.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum RgbColorSpace {
    #[default]
    Srgb,
    DisplayP3,
}

/// Validated normalized red, green, blue, and alpha channels.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct RgbaColor {
    red: f32,
    green: f32,
    blue: f32,
    alpha: f32,
    color_space: RgbColorSpace,
}

impl RgbaColor {
    /// Construct a color whose channels are finite and in the inclusive range 0–1.
    pub fn new(
        red: f32,
        green: f32,
        blue: f32,
        alpha: f32,
        color_space: RgbColorSpace,
    ) -> Result<Self> {
        for (name, value) in [
            ("red", red),
            ("green", green),
            ("blue", blue),
            ("alpha", alpha),
        ] {
            if !value.is_finite() || !(0.0..=1.0).contains(&value) {
                return Err(Error::ParseError(format!(
                    "iWork drawing {name} channel must be finite and between 0 and 1"
                )));
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

    pub const fn red(self) -> f32 {
        self.red
    }

    pub const fn green(self) -> f32 {
        self.green
    }

    pub const fn blue(self) -> f32 {
        self.blue
    }

    pub const fn alpha(self) -> f32 {
        self.alpha
    }

    pub const fn color_space(self) -> RgbColorSpace {
        self.color_space
    }

    pub const fn black() -> Self {
        Self {
            red: 0.0,
            green: 0.0,
            blue: 0.0,
            alpha: 1.0,
            color_space: RgbColorSpace::Srgb,
        }
    }
}

pub(crate) fn color_from_native(color: &tsp::Color) -> Result<RgbaColor> {
    if color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return Err(Error::InvalidFormat(
            "native iWork drawing color is not RGB".to_owned(),
        ));
    }
    let color_space =
        match tsp::color::RgbColorSpace::try_from(color.rgbspace.ok_or_else(|| {
            Error::InvalidFormat("native iWork drawing color has no RGB color space".to_owned())
        })?) {
            Ok(tsp::color::RgbColorSpace::Srgb) => RgbColorSpace::Srgb,
            Ok(tsp::color::RgbColorSpace::P3) => RgbColorSpace::DisplayP3,
            Err(_) => {
                return Err(Error::InvalidFormat(
                    "native iWork drawing color uses an unknown RGB color space".to_owned(),
                ));
            },
        };
    RgbaColor::new(
        color.r.ok_or_else(|| {
            Error::InvalidFormat("native iWork drawing color has no red channel".to_owned())
        })?,
        color.g.ok_or_else(|| {
            Error::InvalidFormat("native iWork drawing color has no green channel".to_owned())
        })?,
        color.b.ok_or_else(|| {
            Error::InvalidFormat("native iWork drawing color has no blue channel".to_owned())
        })?,
        color.a.unwrap_or(1.0),
        color_space,
    )
}

pub(crate) fn color_to_native(color: RgbaColor) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        rgbspace: Some(match color.color_space() {
            RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
            RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
        }),
        a: Some(color.alpha()),
        ..Default::default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn invalid_channels_are_rejected() {
        assert!(RgbaColor::new(-0.1, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
        assert!(RgbaColor::new(0.0, f32::NAN, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
    }

    #[test]
    fn native_rgb_round_trips() {
        let color = RgbaColor::new(0.1, 0.2, 0.3, 0.4, RgbColorSpace::DisplayP3).unwrap();
        assert_eq!(color_from_native(&color_to_native(color)).unwrap(), color);
    }
}
