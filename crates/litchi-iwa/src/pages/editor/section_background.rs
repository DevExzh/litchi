//! Semantic Pages section-background CRUD.

use litchi_iwa_common::color::{RgbColorSpace, Rgba};

use super::*;

const FILL_COLOR_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_RGB_SPACE_FIELD: u32 = 12;

impl PagesEditor {
    /// Read a section background as a semantic solid color when possible.
    pub fn section_background(&self, section_id: u64) -> Result<Background> {
        self.section_background_payload(section_id)?
            .as_deref()
            .map(decode_section_background)
            .transpose()
            .map(|background| background.unwrap_or(Background::None))
    }

    /// Set, clear, or losslessly replace a section background fill.
    ///
    /// Editing an existing solid color patches only changed nested color
    /// scalars, preserving unknown protobuf fields in the fill and color.
    pub fn set_section_background(
        &mut self,
        section_id: u64,
        background: Background,
    ) -> Result<()> {
        validate_section_background(&background)?;
        let current_payload = self.section_background_payload(section_id)?;
        let current = current_payload
            .as_deref()
            .map(decode_section_background)
            .transpose()?
            .unwrap_or(Background::None);
        if current == background {
            return Ok(());
        }

        let payload = match background {
            Background::None => None,
            Background::Opaque(payload) => Some(payload.into_bytes()),
            Background::Solid(color) => {
                let payload = match (current, current_payload.as_deref()) {
                    (Background::Solid(current), Some(payload)) => {
                        patch_solid_background(payload, current, color)?
                    },
                    _ => encode_solid_background(color),
                };
                Some(payload.into_boxed_slice())
            },
            _ => {
                return Err(Error::ParseError(
                    "unsupported Pages section background variant".to_owned(),
                ));
            },
        };
        self.set_section_background_payload(section_id, payload.as_deref())
    }
}

fn validate_section_background(background: &Background) -> Result<()> {
    match background {
        Background::None => Ok(()),
        Background::Solid(color) => validate_pages_color(*color),
        Background::Opaque(payload) => {
            tsd::FillArchive::decode(payload.as_bytes()).map_err(|error| {
                Error::ParseError(format!(
                    "Opaque Pages section background is not a TSD.FillArchive: {error}"
                ))
            })?;
            Ok(())
        },
        _ => Err(Error::ParseError(
            "unsupported Pages section background variant".to_owned(),
        )),
    }
}

fn validate_pages_color(color: Rgba) -> Result<()> {
    Rgba::new(
        color.red(),
        color.green(),
        color.blue(),
        color.alpha(),
        color.color_space(),
    )
    .map(|_| ())
    .map_err(|error| Error::ParseError(format!("invalid Pages section background color: {error}")))
}

fn decode_section_background(payload: &[u8]) -> Result<Background> {
    let fill = tsd::FillArchive::decode(payload)?;
    let Some(color) = fill.color.as_ref() else {
        return Opaque::from_slice(payload)
            .map(Background::Opaque)
            .map_err(|error| Error::ParseError(error.to_string()));
    };
    if fill.gradient.is_some()
        || fill.image.is_some()
        || color.model != tsp::color::ColorModel::Rgb as i32
        || color.c.is_some()
        || color.m.is_some()
        || color.y.is_some()
        || color.k.is_some()
        || color.w.is_some()
    {
        return Opaque::from_slice(payload)
            .map(Background::Opaque)
            .map_err(|error| Error::ParseError(error.to_string()));
    }
    let Some((red, green, blue)) = color
        .r
        .zip(color.g)
        .zip(color.b)
        .map(|((red, green), blue)| (red, green, blue))
    else {
        return Opaque::from_slice(payload)
            .map(Background::Opaque)
            .map_err(|error| Error::ParseError(error.to_string()));
    };
    let color_space = match color.rgbspace {
        None => RgbColorSpace::Srgb,
        Some(value) if value == tsp::color::RgbColorSpace::Srgb as i32 => RgbColorSpace::Srgb,
        Some(value) if value == tsp::color::RgbColorSpace::P3 as i32 => RgbColorSpace::DisplayP3,
        _ => {
            return Opaque::from_slice(payload)
                .map(Background::Opaque)
                .map_err(|error| Error::ParseError(error.to_string()));
        },
    };
    let Ok(semantic) = Rgba::new(red, green, blue, color.a.unwrap_or(1.0), color_space) else {
        return Opaque::from_slice(payload)
            .map(Background::Opaque)
            .map_err(|error| Error::ParseError(error.to_string()));
    };
    Ok(Background::Solid(semantic))
}

fn encode_solid_background(color: Rgba) -> Vec<u8> {
    tsd::FillArchive {
        color: Some(tsp::Color {
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
        }),
        ..Default::default()
    }
    .encode_to_vec()
}

fn patch_solid_background(payload: &[u8], current: Rgba, replacement: Rgba) -> Result<Vec<u8>> {
    let fill = tsd::FillArchive::decode(payload)?;
    let color = fill.color.ok_or_else(|| {
        Error::InvalidFormat("Pages solid section background lost its color".to_owned())
    })?;
    let mut data = payload.to_vec();
    for (field_number, present, before, after) in [
        (
            COLOR_RED_FIELD,
            color.r.is_some(),
            current.red(),
            replacement.red(),
        ),
        (
            COLOR_GREEN_FIELD,
            color.g.is_some(),
            current.green(),
            replacement.green(),
        ),
        (
            COLOR_BLUE_FIELD,
            color.b.is_some(),
            current.blue(),
            replacement.blue(),
        ),
        (
            COLOR_ALPHA_FIELD,
            color.a.is_some(),
            current.alpha(),
            replacement.alpha(),
        ),
    ] {
        if before != after {
            data = patch_nested_fixed32_field(
                &data,
                &[FILL_COLOR_FIELD, field_number],
                present,
                Some(after.to_bits()),
            )?;
        }
    }
    if current.color_space() != replacement.color_space() {
        let rgbspace = match replacement.color_space() {
            RgbColorSpace::Srgb => tsp::color::RgbColorSpace::Srgb as u64,
            RgbColorSpace::DisplayP3 => tsp::color::RgbColorSpace::P3 as u64,
        };
        data = patch_nested_varint_field(
            &data,
            &[FILL_COLOR_FIELD, COLOR_RGB_SPACE_FIELD],
            color.rgbspace.is_some(),
            Some(rgbspace),
        )?;
    }
    if decode_section_background(&data)? != Background::Solid(replacement) {
        return Err(Error::InvalidFormat(
            "Pages solid section background patch failed validation".to_owned(),
        ));
    }
    Ok(data)
}
