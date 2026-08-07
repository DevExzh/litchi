//! Native RGB conversion for Keynote slide-background fills.

use super::*;

use litchi_iwa_common::color::{RgbColorSpace, Rgba};

pub(super) fn color_from_native(color: &tsp::Color) -> Option<Rgba> {
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
        Ok(tsp::color::RgbColorSpace::Srgb) => RgbColorSpace::Srgb,
        Ok(tsp::color::RgbColorSpace::P3) => RgbColorSpace::DisplayP3,
        Err(_) => return None,
    };
    Rgba::new(red, green, blue, color.a.unwrap_or(1.0), color_space).ok()
}

pub(super) fn color_to_native(color: Rgba) -> tsp::Color {
    tsp::Color {
        model: tsp::color::ColorModel::Rgb as i32,
        r: Some(color.red()),
        g: Some(color.green()),
        b: Some(color.blue()),
        rgbspace: Some(native_color_space(color.color_space())),
        a: Some(color.alpha()),
        ..Default::default()
    }
}

pub(super) fn native_color_space(color_space: RgbColorSpace) -> i32 {
    match color_space {
        RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as i32,
        RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as i32,
    }
}
