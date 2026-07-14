//! Typed RGB colors shared by Keynote slide-background fills.

use super::*;

/// RGB color space used by a Keynote slide background.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteRgbColorSpace {
    Srgb,
    DisplayP3,
}

/// Normalized RGBA components used by Keynote's native fill archives.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct KeynoteRgbaColor {
    pub red: f32,
    pub green: f32,
    pub blue: f32,
    pub alpha: f32,
    pub color_space: KeynoteRgbColorSpace,
}

pub(super) fn color_from_native(color: &tsp::Color) -> Option<KeynoteRgbaColor> {
    if color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return None;
    }
    let (Some(red), Some(green), Some(blue), Some(rgbspace)) =
        (color.r, color.g, color.b, color.rgbspace)
    else {
        return None;
    };
    let color_space = match tsp::color::RgbColorSpace::try_from(rgbspace) {
        Ok(tsp::color::RgbColorSpace::Srgb) => KeynoteRgbColorSpace::Srgb,
        Ok(tsp::color::RgbColorSpace::P3) => KeynoteRgbColorSpace::DisplayP3,
        Err(_) => return None,
    };
    Some(KeynoteRgbaColor {
        red,
        green,
        blue,
        alpha: color.a.unwrap_or(1.0),
        color_space,
    })
}

pub(super) fn color_to_native(color: KeynoteRgbaColor) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(color.red),
        g: Some(color.green),
        b: Some(color.blue),
        rgbspace: Some(native_color_space(color.color_space)),
        a: Some(color.alpha),
        ..Default::default()
    }
}

pub(super) fn validate_color(color: KeynoteRgbaColor, context: &str) -> Result<()> {
    for (name, value) in [
        ("red", color.red),
        ("green", color.green),
        ("blue", color.blue),
        ("alpha", color.alpha),
    ] {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err(Error::ParseError(format!(
                "{context} {name} channel must be finite and between 0 and 1"
            )));
        }
    }
    Ok(())
}

pub(super) fn native_color_space(color_space: KeynoteRgbColorSpace) -> i32 {
    match color_space {
        KeynoteRgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
        KeynoteRgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
    }
}
