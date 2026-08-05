//! Native RGB conversion for the shared validated color value.

use crate::protobuf::tsp;
use crate::{Error, Result};

use litchi_iwa_common::color::Error as ColorError;
pub use litchi_iwa_common::color::{RgbColorSpace, Rgba};

/// The facade spelling for the shared RGBA value.
pub type RgbaColor = Rgba;

impl From<ColorError> for crate::Error {
    fn from(error: ColorError) -> Self {
        Self::ParseError(error.to_string())
    }
}

pub(crate) fn color_from_native(color: &tsp::Color) -> Result<Rgba> {
    if color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return Err(Error::InvalidFormat(
            "native iWork color is not RGB".to_owned(),
        ));
    }
    let color_space =
        match tsp::color::RgbColorSpace::try_from(color.rgbspace.ok_or_else(|| {
            Error::InvalidFormat("native iWork color has no RGB color space".to_owned())
        })?) {
            Ok(tsp::color::RgbColorSpace::Srgb) => RgbColorSpace::Srgb,
            Ok(tsp::color::RgbColorSpace::P3) => RgbColorSpace::DisplayP3,
            Err(_) => {
                return Err(Error::InvalidFormat(
                    "native iWork color uses an unknown RGB color space".to_owned(),
                ));
            },
        };
    Ok(Rgba::new(
        color.r.ok_or_else(|| {
            Error::InvalidFormat("native iWork color has no red channel".to_owned())
        })?,
        color.g.ok_or_else(|| {
            Error::InvalidFormat("native iWork color has no green channel".to_owned())
        })?,
        color.b.ok_or_else(|| {
            Error::InvalidFormat("native iWork color has no blue channel".to_owned())
        })?,
        color.a.unwrap_or(1.0),
        color_space,
    )?)
}

pub(crate) fn color_to_native(color: Rgba) -> tsp::Color {
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
        assert!(Rgba::new(-0.1, 0.0, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
        assert!(Rgba::new(0.0, f32::NAN, 0.0, 1.0, RgbColorSpace::Srgb).is_err());
    }

    #[test]
    fn native_rgb_round_trips() {
        let color = Rgba::new(0.1, 0.2, 0.3, 0.4, RgbColorSpace::DisplayP3).unwrap();
        assert_eq!(color_from_native(&color_to_native(color)).unwrap(), color);
    }
}
