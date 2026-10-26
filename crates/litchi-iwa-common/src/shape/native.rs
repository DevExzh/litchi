//! The semantic reading of stored native drawing values: colors, strokes,
//! gradients and image fills as `litchi_iwa_protos::shape_style_codec::native`
//! decodes them (plain native fields), converted to the shared values with
//! the native rules — defaults, enum values, unsupported forms (record 0375).

use crate::color::{RgbColorSpace, Rgba};
use crate::shape::fill::{
    self, Angle, Gradient, Image, ImageData, ImageTechnique, Kind, Opacity, Stop, StopMidpoint,
    StopPosition,
};
use crate::shape::geometry::Size;
use crate::shape::stroke::{Cap, Join, MiterLimit, Pattern, Stroke, Width};

const RGB_MODEL: i32 = 1;
const SRGB_SPACE: i32 = 1;
const P3_SPACE: i32 = 2;
const DEFAULT_MITER_LIMIT: f32 = 4.0;
const SHORT_DASH: [f32; 2] = [1.0, 1.0];
const MEDIUM_DASH: [f32; 2] = [2.0, 2.0];
const LONG_DASH: [f32; 2] = [6.0, 6.0];
const ROUNDED_DASH: [f32; 2] = [0.001, 2.0];
const PATTERN_KIND_DASH: i32 = 0;
const PATTERN_KIND_SOLID: i32 = 1;
const PATTERN_KIND_EMPTY: i32 = 2;

/// Why a stored native value has no shared semantic value.
#[derive(Debug, Clone, Copy, PartialEq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The color is not an RGB color (or carries CMYK/white channels).
    #[error("native color is not RGB")]
    NotRgb,
    /// A required native field is absent.
    #[error("native {0} is missing")]
    Missing(&'static str),
    /// An enumeration holds an unknown value.
    #[error("native {what} has the unknown value {value}")]
    Unknown {
        /// The enumeration.
        what: &'static str,
        /// The stored value.
        value: i32,
    },
    /// The stored form is valid but not a standard shared value.
    #[error("native {0} is not a standard value")]
    Unsupported(&'static str),
    /// A color channel is invalid.
    #[error(transparent)]
    Color(#[from] crate::color::Error),
    /// A fill value is invalid.
    #[error(transparent)]
    Fill(#[from] fill::Error),
    /// A stroke value is invalid.
    #[error(transparent)]
    Stroke(#[from] crate::shape::stroke::Error),
}

/// Result type of the native conversions.
pub type Result<T> = std::result::Result<T, Error>;

/// A stored `TSP.Color`: model, RGB space, `[r, g, b, a]` and whether any
/// CMYK/white channel is present. The alpha defaults to 1.
///
/// # Errors
///
/// [`Error::NotRgb`], [`Error::Missing`], [`Error::Unknown`] or an invalid
/// channel.
pub fn color(
    model: Option<i32>,
    rgb_space: Option<i32>,
    rgba: [Option<f32>; 4],
    has_other_channels: bool,
) -> Result<Rgba> {
    if model != Some(RGB_MODEL) || has_other_channels {
        return Err(Error::NotRgb);
    }
    let space = match rgb_space.ok_or(Error::Missing("color space"))? {
        SRGB_SPACE => RgbColorSpace::Srgb,
        P3_SPACE => RgbColorSpace::DisplayP3,
        value => {
            return Err(Error::Unknown {
                what: "color space",
                value,
            });
        },
    };
    let [red, green, blue, alpha] = rgba;
    Ok(Rgba::new(
        red.ok_or(Error::Missing("red channel"))?,
        green.ok_or(Error::Missing("green channel"))?,
        blue.ok_or(Error::Missing("blue channel"))?,
        alpha.unwrap_or(1.0),
        space,
    )?)
}

/// A stored stroke pattern: `type`, `count` and the lengths.
#[derive(Debug, Clone, Copy)]
pub struct StoredPattern<'a> {
    /// `type` (absent reads as the dash pattern, 0).
    pub kind: Option<i32>,
    /// `count`.
    pub count: Option<u32>,
    /// The pattern lengths.
    pub lengths: &'a [f32],
}

/// A stored `TSD.StrokeArchive`, `None` for the empty pattern (a hidden
/// stroke). The color is the already converted stored color, whose error is
/// reported only when the stroke uses it.
///
/// # Errors
///
/// [`Error::Unsupported`] for smart, frame and patterned strokes and
/// non-standard dashes, [`Error::Missing`]/[`Error::Unknown`] for invalid
/// fields, or an invalid width or miter limit.
#[allow(clippy::too_many_arguments, reason = "one argument per stored field")]
pub fn stroke(
    color: Option<Result<Rgba>>,
    width: Option<f32>,
    cap: Option<i32>,
    join: Option<i32>,
    miter_limit: Option<f32>,
    pattern: Option<StoredPattern<'_>>,
    is_special: bool,
) -> Result<Option<Stroke>> {
    if is_special {
        return Err(Error::Unsupported("smart, frame or patterned stroke"));
    }
    let pattern = pattern.ok_or(Error::Missing("stroke pattern"))?;
    let kind = pattern.kind.unwrap_or(PATTERN_KIND_DASH);
    let typed = match kind {
        PATTERN_KIND_EMPTY => return Ok(None),
        PATTERN_KIND_SOLID => Pattern::Solid,
        PATTERN_KIND_DASH => {
            let used = pattern.count.unwrap_or_default() as usize;
            let values = pattern
                .lengths
                .get(..used)
                .ok_or(Error::Unsupported("truncated stroke pattern"))?;
            match values {
                [a, b] if [*a, *b] == SHORT_DASH => Pattern::ShortDash,
                [a, b] if [*a, *b] == MEDIUM_DASH => Pattern::MediumDash,
                [a, b] if [*a, *b] == LONG_DASH => Pattern::LongDash,
                [a, b] if [*a, *b] == ROUNDED_DASH => Pattern::RoundedDash,
                _ => return Err(Error::Unsupported("stroke dash pattern")),
            }
        },
        value => {
            return Err(Error::Unknown {
                what: "stroke pattern type",
                value,
            });
        },
    };
    let color = color.ok_or(Error::Missing("stroke color"))??;
    let width = Width::new(width.ok_or(Error::Missing("stroke width"))?)?;
    let cap = match cap.unwrap_or_default() {
        0 => Cap::Butt,
        1 => Cap::Round,
        2 => Cap::Square,
        value => {
            return Err(Error::Unknown {
                what: "stroke cap",
                value,
            });
        },
    };
    let join = match join.unwrap_or_default() {
        0 => Join::Miter,
        1 => Join::Round,
        2 => Join::Bevel,
        value => {
            return Err(Error::Unknown {
                what: "stroke join",
                value,
            });
        },
    };
    let miter_limit = MiterLimit::new(miter_limit.unwrap_or(DEFAULT_MITER_LIMIT))?;
    Ok(Some(Stroke {
        color,
        width,
        pattern: typed,
        cap,
        join,
        miter_limit,
    }))
}

/// A stored gradient stop.
///
/// # Errors
///
/// [`Error::Missing`] or an invalid position or midpoint.
pub fn stop(color: Option<Rgba>, fraction: Option<f32>, inflection: Option<f32>) -> Result<Stop> {
    Ok(Stop::new(
        color.ok_or(Error::Missing("gradient stop color"))?,
        StopPosition::new(fraction.ok_or(Error::Missing("gradient stop position"))?)?,
        StopMidpoint::new(inflection.ok_or(Error::Missing("gradient stop midpoint"))?)?,
    ))
}

/// A stored `TSD.GradientArchive` (angle in radians).
///
/// # Errors
///
/// [`Error::Unsupported`] for a transformed gradient, [`Error::Missing`] or
/// [`Error::Unknown`] for invalid fields, or an invalid gradient.
pub fn gradient(
    kind: Option<i32>,
    stops: Vec<Stop>,
    opacity: Option<f32>,
    advanced: Option<bool>,
    angle: Option<f32>,
    is_transformed: bool,
) -> Result<Gradient> {
    if is_transformed {
        return Err(Error::Unsupported("transformed gradient"));
    }
    let kind = match kind.ok_or(Error::Missing("gradient type"))? {
        0 => Kind::Linear,
        1 => Kind::Radial,
        value => {
            return Err(Error::Unknown {
                what: "gradient type",
                value,
            });
        },
    };
    let angle = angle.ok_or(Error::Missing("gradient angle"))?;
    Ok(Gradient::from_parts(
        kind,
        stops,
        Opacity::new(opacity.ok_or(Error::Missing("gradient opacity"))?)?,
        advanced.ok_or(Error::Missing("gradient editing mode"))?,
        Angle::from_radians(angle)?,
    )?)
}

/// A stored `TSD.ImageFillArchive` (`fill_size` as width, height; `data` 0
/// reads as the placeholder).
///
/// # Errors
///
/// [`Error::Unsupported`] for legacy or database-backed images,
/// [`Error::Missing`]/[`Error::Unknown`] for invalid fields, or an invalid
/// size.
#[allow(clippy::too_many_arguments, reason = "one argument per stored field")]
pub fn image(
    data: Option<u64>,
    technique: Option<i32>,
    fill_size: Option<[f32; 2]>,
    tint: Option<Rgba>,
    reference_color: Option<Rgba>,
    generic: Option<bool>,
    is_legacy: bool,
) -> Result<Image> {
    if is_legacy {
        return Err(Error::Unsupported("legacy or database-backed image fill"));
    }
    let technique = technique.unwrap_or_default();
    let technique = ImageTechnique::from_native(technique).ok_or(Error::Unknown {
        what: "image fill technique",
        value: technique,
    })?;
    let [width, height] = fill_size.ok_or(Error::Missing("image fill size"))?;
    let size = Size { width, height };
    let mut image = match data.filter(|identifier| *identifier != 0) {
        Some(identifier) => Image::embedded(ImageData::new(identifier)?, technique, size)?,
        None => Image::placeholder(technique, size)?,
    };
    if let Some(tint) = tint {
        image = image.with_tint(tint);
    }
    if let Some(reference) = reference_color {
        image = image.with_reference_color(reference);
    }
    Ok(image.with_generic_untagged_image(generic.unwrap_or(false)))
}

/// A stored `TSD.FillArchive` from its at-most-one representation.
///
/// # Errors
///
/// [`Error::Unsupported`] when more than one representation is present.
pub fn fill(
    color: Option<Rgba>,
    gradient: Option<Gradient>,
    image: Option<Image>,
) -> Result<fill::Fill> {
    match (color, gradient, image) {
        (None, None, None) => Ok(fill::Fill::None),
        (Some(color), None, None) => Ok(fill::Fill::Solid(color)),
        (None, Some(gradient), None) => Ok(fill::Fill::Gradient(gradient)),
        (None, None, Some(image)) => Ok(fill::Fill::Image(image)),
        _ => Err(Error::Unsupported("fill with several representations")),
    }
}

// ---- the stored form of the shared values (the encode direction) -------

/// The stored fields of a `TSP.Color`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StoredColor {
    /// `model` (RGB).
    pub model: i32,
    /// `r`, `g`, `b`, `a`.
    pub rgba: [f32; 4],
    /// `rgbspace`.
    pub rgb_space: i32,
}

/// The stored fields of a `TSD.StrokeArchive` (pattern with its six native
/// slots).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StoredStroke {
    /// `color`.
    pub color: StoredColor,
    /// `width`.
    pub width: f32,
    /// `cap`.
    pub cap: i32,
    /// `join`.
    pub join: i32,
    /// `miter_limit`.
    pub miter_limit: f32,
    /// `pattern.type`.
    pub pattern_kind: i32,
    /// `pattern.count`.
    pub pattern_count: u32,
    /// `pattern.pattern` (always six slots).
    pub pattern_lengths: [f32; 6],
}

/// The stored fields of a gradient stop.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StoredStop {
    /// `color`.
    pub color: StoredColor,
    /// `fraction`.
    pub fraction: f32,
    /// `inflection`.
    pub inflection: f32,
}

/// The stored fields of a `TSD.GradientArchive` (angle in radians).
#[derive(Debug, Clone, PartialEq)]
pub struct StoredGradient {
    /// `type`.
    pub kind: i32,
    /// `stops`.
    pub stops: Vec<StoredStop>,
    /// `opacity`.
    pub opacity: f32,
    /// `advancedGradient`.
    pub advanced: bool,
    /// `anglegradient.gradientangle`.
    pub angle: f32,
}

/// The stored fields of a `TSD.ImageFillArchive` (`data` 0 for the
/// placeholder).
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct StoredImage {
    /// `imagedata.identifier`.
    pub data: u64,
    /// `technique`.
    pub technique: i32,
    /// `tint`.
    pub tint: Option<StoredColor>,
    /// `fillsize`.
    pub fill_size: [f32; 2],
    /// `interpretsUntaggedImageDataAsGeneric`.
    pub generic: bool,
    /// `referencecolor`.
    pub reference_color: Option<StoredColor>,
}

/// The stored fields of a `TSD.FillArchive`.
#[derive(Debug, Clone, PartialEq, Default)]
pub struct StoredFill {
    /// `color`.
    pub color: Option<StoredColor>,
    /// `gradient`.
    pub gradient: Option<StoredGradient>,
    /// `image`.
    pub image: Option<StoredImage>,
}

/// The stored form of a color.
#[must_use]
pub fn stored_color(color: Rgba) -> StoredColor {
    StoredColor {
        model: RGB_MODEL,
        rgba: [color.red(), color.green(), color.blue(), color.alpha()],
        rgb_space: match color.color_space() {
            RgbColorSpace::Srgb => SRGB_SPACE,
            RgbColorSpace::DisplayP3 => P3_SPACE,
        },
    }
}

/// The stored form of a stroke; `None` stores the native hidden stroke (a
/// one-point black solid stroke with the empty pattern).
#[must_use]
pub fn stored_stroke(stroke: Option<Stroke>) -> StoredStroke {
    let (stroke, hidden) = match stroke {
        Some(stroke) => (stroke, false),
        None => (Stroke::new(Rgba::black(), Width::ONE, Pattern::Solid), true),
    };
    let (kind, count, lengths) = match stroke.pattern {
        Pattern::Solid => (PATTERN_KIND_SOLID, 0, [0.0, 0.0]),
        Pattern::ShortDash => (PATTERN_KIND_DASH, 2, SHORT_DASH),
        Pattern::MediumDash => (PATTERN_KIND_DASH, 2, MEDIUM_DASH),
        Pattern::LongDash => (PATTERN_KIND_DASH, 2, LONG_DASH),
        Pattern::RoundedDash => (PATTERN_KIND_DASH, 2, ROUNDED_DASH),
    };
    let (kind, count, lengths) = if hidden {
        (PATTERN_KIND_EMPTY, 0, [0.0, 0.0])
    } else {
        (kind, count, lengths)
    };
    StoredStroke {
        color: stored_color(stroke.color),
        width: stroke.width.points(),
        cap: match stroke.cap {
            Cap::Butt => 0,
            Cap::Round => 1,
            Cap::Square => 2,
        },
        join: match stroke.join {
            Join::Miter => 0,
            Join::Round => 1,
            Join::Bevel => 2,
        },
        miter_limit: stroke.miter_limit.ratio(),
        pattern_kind: kind,
        pattern_count: count,
        pattern_lengths: [lengths[0], lengths[1], 0.0, 0.0, 0.0, 0.0],
    }
}

/// The stored form of a fill.
#[must_use]
pub fn stored_fill(fill: &fill::Fill) -> StoredFill {
    match fill {
        fill::Fill::None => StoredFill::default(),
        fill::Fill::Solid(color) => StoredFill {
            color: Some(stored_color(*color)),
            ..StoredFill::default()
        },
        fill::Fill::Gradient(gradient) => StoredFill {
            gradient: Some(StoredGradient {
                kind: match gradient.kind() {
                    Kind::Linear => 0,
                    Kind::Radial => 1,
                },
                stops: gradient
                    .stops()
                    .iter()
                    .map(|stop| StoredStop {
                        color: stored_color(stop.color()),
                        fraction: stop.position().get(),
                        inflection: stop.midpoint().get(),
                    })
                    .collect(),
                opacity: gradient.opacity().get(),
                advanced: gradient.is_advanced(),
                angle: gradient.angle().radians(),
            }),
            ..StoredFill::default()
        },
        fill::Fill::Image(image) => StoredFill {
            image: Some(StoredImage {
                data: image.data().map_or(0, ImageData::get),
                technique: image.technique().native_value(),
                tint: image.tint().map(stored_color),
                fill_size: [image.fill_size().width, image.fill_size().height],
                generic: image.interprets_untagged_image_as_generic(),
                reference_color: image.reference_color().map(stored_color),
            }),
            ..StoredFill::default()
        },
    }
}

#[cfg(test)]
mod tests {
    use super::{Error, StoredPattern, color, fill, image, stroke};
    use crate::shape::fill::{Fill, ImageTechnique};
    use crate::shape::stroke::Pattern;

    #[test]
    fn native_values_convert_with_native_rules() {
        let red = color(
            Some(1),
            Some(1),
            [Some(1.0), Some(0.0), Some(0.0), None],
            false,
        )
        .unwrap();
        assert_eq!(red.alpha(), 1.0);
        assert_eq!(
            color(Some(2), Some(1), [None; 4], false),
            Err(Error::NotRgb)
        );
        let dash = StoredPattern {
            kind: Some(0),
            count: Some(2),
            lengths: &[2.0, 2.0, 0.0, 0.0],
        };
        let stroke_value = stroke(
            Some(Ok(red)),
            Some(2.0),
            None,
            None,
            None,
            Some(dash),
            false,
        )
        .unwrap()
        .unwrap();
        assert_eq!(stroke_value.pattern, Pattern::MediumDash);
        let empty = StoredPattern {
            kind: Some(2),
            count: Some(0),
            lengths: &[],
        };
        assert_eq!(
            stroke(None, None, None, None, None, Some(empty), false),
            Ok(None)
        );
        let placeholder = image(
            Some(0),
            Some(4),
            Some([10.0, 10.0]),
            None,
            None,
            None,
            false,
        )
        .unwrap();
        assert_eq!(placeholder.data(), None);
        assert_eq!(placeholder.technique(), ImageTechnique::ScaleToFit);
        assert_eq!(fill(None, None, None), Ok(Fill::None));
        let hidden = super::stored_stroke(None);
        assert_eq!((hidden.pattern_kind, hidden.width), (2, 1.0));
        let stored = super::stored_stroke(Some(stroke_value));
        assert_eq!((stored.pattern_kind, stored.pattern_lengths[0]), (0, 2.0));
        assert_eq!(super::stored_color(red).rgb_space, 1);
        assert!(fill(Some(red), None, Some(placeholder)).is_err());
    }
}
