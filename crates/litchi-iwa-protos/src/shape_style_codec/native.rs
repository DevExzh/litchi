//! Strict borrowed decoders of the native drawing values shared by shape and
//! chart styles: `TSP.Color`, `TSD.StrokeArchive`, `TSD.FillArchive` (color,
//! gradient and image fills) and the shape-style property slots of
//! `TSWP.ShapeStyleArchive`. The results are plain native fields; the
//! semantic conversion lives in `litchi_iwa_common::shape::native`
//! (record 0375).

use crate::chart_style_codec::{
    Budget, Error, Field, Limits, check_bytes, read_f32, require_wire_type, singular, visit_fields,
};

const WIRE_VARINT: u8 = 0;
const WIRE_LENGTH_DELIMITED: u8 = 2;
const WIRE_FIXED32: u8 = 5;

fn optional_float(data: &[u8], number: u32, budget: &mut Budget) -> Result<Option<f32>, Error> {
    let Some(field) = singular(data, number, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_FIXED32)?;
    Ok(Some(read_f32(field.payload(data))))
}

fn optional_varint(data: &[u8], number: u32, budget: &mut Budget) -> Result<Option<u64>, Error> {
    let Some(field) = singular(data, number, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_VARINT)?;
    Ok(Some(field.varint))
}

fn optional_message<'a>(
    data: &'a [u8],
    number: u32,
    budget: &mut Budget,
) -> Result<Option<&'a [u8]>, Error> {
    let Some(field) = singular(data, number, budget)? else {
        return Ok(None);
    };
    require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
    Ok(Some(field.payload(data)))
}

fn repeated(data: &[u8], number: u32, budget: &mut Budget) -> Result<Vec<Field>, Error> {
    let mut fields = Vec::new();
    visit_fields(data, budget, |field| {
        if field.number == number {
            fields.push(field);
        }
        Ok(())
    })?;
    Ok(fields)
}

/// A `TSP.Color` as stored.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Color {
    /// `model` (1 = RGB).
    pub model: Option<i32>,
    /// `r`, `g`, `b`, `a`.
    pub rgba: [Option<f32>; 4],
    /// `rgbspace` (1 = sRGB, 2 = Display P3).
    pub rgb_space: Option<i32>,
    /// Whether any CMYK or white channel is present.
    pub has_other_channels: bool,
}

fn color_at(data: &[u8], budget: &mut Budget) -> Result<Color, Error> {
    let mut other = false;
    for number in [7, 8, 9, 10, 11] {
        other |= optional_float(data, number, budget)?.is_some();
    }
    Ok(Color {
        model: optional_varint(data, 1, budget)?.map(|value| value as i32),
        rgba: [
            optional_float(data, 3, budget)?,
            optional_float(data, 4, budget)?,
            optional_float(data, 5, budget)?,
            optional_float(data, 6, budget)?,
        ],
        rgb_space: optional_varint(data, 12, budget)?.map(|value| value as i32),
        has_other_channels: other,
    })
}

/// Decode a `TSP.Color` payload.
///
/// # Errors
///
/// Framing, duplicate-field, wire-type and limit errors.
pub fn color(data: &[u8], limits: Limits) -> Result<Color, Error> {
    check_bytes(data.len(), limits)?;
    color_at(data, &mut Budget::new(limits))
}

/// A `TSD.StrokePatternArchive` as stored.
#[derive(Debug, Clone, PartialEq)]
pub struct StrokePattern {
    /// `type` (0 = pattern, 1 = solid, 2 = empty).
    pub kind: Option<i32>,
    /// `phase`.
    pub phase: Option<f32>,
    /// `count`.
    pub count: Option<u32>,
    /// `pattern` lengths.
    pub lengths: Vec<f32>,
}

/// A `TSD.StrokeArchive` as stored.
#[derive(Debug, Clone, PartialEq)]
pub struct Stroke {
    /// `color`.
    pub color: Option<Color>,
    /// `width`.
    pub width: Option<f32>,
    /// `cap`.
    pub cap: Option<i32>,
    /// `join`.
    pub join: Option<i32>,
    /// `miter_limit`.
    pub miter_limit: Option<f32>,
    /// `pattern`.
    pub pattern: Option<StrokePattern>,
    /// Whether a smart, frame or patterned stroke is present.
    pub is_special: bool,
}

/// Decode a `TSD.StrokeArchive` payload.
///
/// # Errors
///
/// Framing, duplicate-field, wire-type and limit errors.
pub fn stroke(data: &[u8], limits: Limits) -> Result<Stroke, Error> {
    check_bytes(data.len(), limits)?;
    let budget = &mut Budget::new(limits);
    let color = optional_message(data, 1, budget)?
        .map(|color| color_at(color, budget))
        .transpose()?;
    let pattern = match optional_message(data, 6, budget)? {
        Some(pattern) => {
            let mut lengths = Vec::new();
            for field in repeated(pattern, 4, budget)? {
                match field.wire_type {
                    WIRE_FIXED32 => lengths.push(read_f32(field.payload(pattern))),
                    WIRE_LENGTH_DELIMITED => {
                        let packed = field.payload(pattern);
                        if packed.len() % 4 != 0 {
                            return Err(Error::Malformed {
                                reason: "a packed stroke pattern is not a whole number of floats",
                            });
                        }
                        lengths.extend(packed.chunks_exact(4).map(read_f32));
                    },
                    found => {
                        return Err(Error::WireType {
                            field: 4,
                            expected: WIRE_FIXED32,
                            found,
                        });
                    },
                }
            }
            Some(StrokePattern {
                kind: optional_varint(pattern, 1, budget)?.map(|value| value as i32),
                phase: optional_float(pattern, 2, budget)?,
                count: optional_varint(pattern, 3, budget)?.map(|value| value as u32),
                lengths,
            })
        },
        None => None,
    };
    let mut is_special = false;
    for number in [7, 8, 9] {
        is_special |= optional_message(data, number, budget)?.is_some();
    }
    Ok(Stroke {
        color,
        width: optional_float(data, 2, budget)?,
        cap: optional_varint(data, 3, budget)?.map(|value| value as i32),
        join: optional_varint(data, 4, budget)?.map(|value| value as i32),
        miter_limit: optional_float(data, 5, budget)?,
        pattern,
        is_special,
    })
}

/// One `TSD.GradientArchive.GradientStop`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientStop {
    /// `color`.
    pub color: Option<Color>,
    /// `fraction`.
    pub fraction: Option<f32>,
    /// `inflection`.
    pub inflection: Option<f32>,
}

/// A `TSD.GradientArchive` as stored.
#[derive(Debug, Clone, PartialEq)]
pub struct Gradient {
    /// `type` (0 = linear, 1 = radial).
    pub kind: Option<i32>,
    /// `stops`.
    pub stops: Vec<GradientStop>,
    /// `opacity`.
    pub opacity: Option<f32>,
    /// `advancedGradient`.
    pub advanced: Option<bool>,
    /// `anglegradient.gradientangle` (radians).
    pub angle: Option<f32>,
    /// Whether a `transformgradient` is present.
    pub is_transformed: bool,
}

/// A `TSD.ImageFillArchive` as stored.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ImageFill {
    /// `imagedata.identifier` (0 reads as absent).
    pub data: Option<u64>,
    /// `technique`.
    pub technique: Option<i32>,
    /// `tint`.
    pub tint: Option<Color>,
    /// `fillsize` (width, height).
    pub fill_size: Option<[f32; 2]>,
    /// `interpretsUntaggedImageDataAsGeneric`.
    pub generic: Option<bool>,
    /// `referencecolor`.
    pub reference_color: Option<Color>,
    /// Whether a legacy or database-backed image reference is present.
    pub is_legacy: bool,
}

/// A `TSD.FillArchive` as stored.
#[derive(Debug, Clone, PartialEq)]
pub struct Fill {
    /// `color`.
    pub color: Option<Color>,
    /// `gradient`.
    pub gradient: Option<Gradient>,
    /// `image`.
    pub image: Option<ImageFill>,
}

fn pair(data: &[u8], budget: &mut Budget) -> Result<[f32; 2], Error> {
    Ok([
        optional_float(data, 1, budget)?.unwrap_or(0.0),
        optional_float(data, 2, budget)?.unwrap_or(0.0),
    ])
}

/// Decode a `TSD.FillArchive` payload.
///
/// # Errors
///
/// Framing, duplicate-field, wire-type and limit errors.
pub fn fill(data: &[u8], limits: Limits) -> Result<Fill, Error> {
    check_bytes(data.len(), limits)?;
    let budget = &mut Budget::new(limits);
    let color = optional_message(data, 1, budget)?
        .map(|color| color_at(color, budget))
        .transpose()?;
    let gradient = match optional_message(data, 2, budget)? {
        Some(gradient) => {
            let mut stops = Vec::new();
            for field in repeated(gradient, 2, budget)? {
                require_wire_type(&field, WIRE_LENGTH_DELIMITED)?;
                let stop = field.payload(gradient);
                stops.push(GradientStop {
                    color: optional_message(stop, 1, budget)?
                        .map(|color| color_at(color, budget))
                        .transpose()?,
                    fraction: optional_float(stop, 2, budget)?,
                    inflection: optional_float(stop, 3, budget)?,
                });
            }
            let angle = match optional_message(gradient, 5, budget)? {
                Some(angle) => optional_float(angle, 2, budget)?,
                None => None,
            };
            Some(Gradient {
                kind: optional_varint(gradient, 1, budget)?.map(|value| value as i32),
                stops,
                opacity: optional_float(gradient, 3, budget)?,
                advanced: optional_varint(gradient, 4, budget)?.map(|value| value != 0),
                angle,
                is_transformed: optional_message(gradient, 6, budget)?.is_some(),
            })
        },
        None => None,
    };
    let image = match optional_message(data, 3, budget)? {
        Some(image) => {
            let data_reference = match optional_message(image, 6, budget)? {
                Some(reference) => optional_varint(reference, 1, budget)?.filter(|id| *id != 0),
                None => None,
            };
            let mut is_legacy = false;
            for number in [1, 5, 7] {
                is_legacy |= optional_message(image, number, budget)?.is_some();
            }
            Some(ImageFill {
                data: data_reference,
                technique: optional_varint(image, 2, budget)?.map(|value| value as i32),
                tint: optional_message(image, 3, budget)?
                    .map(|color| color_at(color, budget))
                    .transpose()?,
                fill_size: optional_message(image, 4, budget)?
                    .map(|size| pair(size, budget))
                    .transpose()?,
                generic: optional_varint(image, 8, budget)?.map(|value| value != 0),
                reference_color: optional_message(image, 9, budget)?
                    .map(|color| color_at(color, budget))
                    .transpose()?,
                is_legacy,
            })
        },
        None => None,
    };
    Ok(Fill {
        color,
        gradient,
        image,
    })
}

/// The inheritance and drawing-property slots of one
/// `TSWP.ShapeStyleArchive` (message type 2025).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeStyle<'source> {
    /// `TSS.StyleArchive.parent` (0 reads as absent).
    pub parent: Option<u64>,
    /// The `TSD.ShapeStylePropertiesArchive.fill` payload, if set here.
    pub fill: Option<&'source [u8]>,
    /// The `TSD.ShapeStylePropertiesArchive.stroke` payload, if set here.
    pub stroke: Option<&'source [u8]>,
}

/// Decode the inheritance and drawing-property slots of a shape style.
///
/// # Errors
///
/// Framing, duplicate-field, wire-type and limit errors, and
/// [`Error::Malformed`] when the required `TSD.ShapeStyleArchive` is absent.
pub fn shape_style(data: &[u8], limits: Limits) -> Result<ShapeStyle<'_>, Error> {
    check_bytes(data.len(), limits)?;
    let budget = &mut Budget::new(limits);
    let drawing = optional_message(data, 1, budget)?.ok_or(Error::Malformed {
        reason: "the shape style has no drawing style",
    })?;
    let parent = match optional_message(drawing, 1, budget)? {
        Some(style) => match optional_message(style, 3, budget)? {
            Some(reference) => optional_varint(reference, 1, budget)?.filter(|id| *id != 0),
            None => None,
        },
        None => None,
    };
    let (fill, stroke) = match optional_message(drawing, 11, budget)? {
        Some(properties) => (
            optional_message(properties, 1, budget)?,
            optional_message(properties, 2, budget)?,
        ),
        None => (None, None),
    };
    Ok(ShapeStyle {
        parent,
        fill,
        stroke,
    })
}

// ---- encoders -----------------------------------------------------------
//
// Each encoder writes the fields a protobuf encoder writes for the stored
// value, in field-number order: required fields always, optional fields when
// present, repeated scalars unpacked (proto2).

fn key(output: &mut Vec<u8>, number: u32, wire_type: u8) {
    crate::chart_style_codec::push_varint(output, (u64::from(number) << 3) | u64::from(wire_type));
}

fn put_float(output: &mut Vec<u8>, number: u32, value: Option<f32>) {
    if let Some(value) = value {
        key(output, number, WIRE_FIXED32);
        output.extend_from_slice(&value.to_le_bytes());
    }
}

fn put_varint(output: &mut Vec<u8>, number: u32, value: Option<u64>) {
    if let Some(value) = value {
        key(output, number, WIRE_VARINT);
        crate::chart_style_codec::push_varint(output, value);
    }
}

fn put_enum(output: &mut Vec<u8>, number: u32, value: Option<i32>) {
    put_varint(output, number, value.map(|value| i64::from(value) as u64));
}

fn put_message(output: &mut Vec<u8>, number: u32, message: &[u8]) {
    key(output, number, WIRE_LENGTH_DELIMITED);
    crate::chart_style_codec::push_varint(output, message.len() as u64);
    output.extend_from_slice(message);
}

/// Encode a `TSP.Color` (`model` is required: absent encodes as 0).
#[must_use]
pub fn encode_color(value: &Color) -> Vec<u8> {
    let mut output = Vec::with_capacity(32);
    put_enum(&mut output, 1, Some(value.model.unwrap_or(0)));
    for (number, channel) in (3..).zip(value.rgba) {
        put_float(&mut output, number, channel);
    }
    put_enum(&mut output, 12, value.rgb_space);
    output
}

/// Encode a `TSD.StrokeArchive` (smart, frame and patterned strokes are not
/// encodable: `is_special` is ignored).
#[must_use]
pub fn encode_stroke(value: &Stroke) -> Vec<u8> {
    let mut output = Vec::with_capacity(80);
    if let Some(color) = &value.color {
        put_message(&mut output, 1, &encode_color(color));
    }
    put_float(&mut output, 2, value.width);
    put_enum(&mut output, 3, value.cap);
    put_enum(&mut output, 4, value.join);
    put_float(&mut output, 5, value.miter_limit);
    if let Some(pattern) = &value.pattern {
        let mut nested = Vec::with_capacity(40);
        put_enum(&mut nested, 1, pattern.kind);
        put_float(&mut nested, 2, pattern.phase);
        put_varint(&mut nested, 3, pattern.count.map(u64::from));
        for length in &pattern.lengths {
            put_float(&mut nested, 4, Some(*length));
        }
        put_message(&mut output, 6, &nested);
    }
    output
}

fn encode_gradient(value: &Gradient) -> Vec<u8> {
    let mut output = Vec::with_capacity(96);
    put_enum(&mut output, 1, value.kind);
    for stop in &value.stops {
        let mut nested = Vec::with_capacity(48);
        if let Some(color) = &stop.color {
            put_message(&mut nested, 1, &encode_color(color));
        }
        put_float(&mut nested, 2, stop.fraction);
        put_float(&mut nested, 3, stop.inflection);
        put_message(&mut output, 2, &nested);
    }
    put_float(&mut output, 3, value.opacity);
    put_varint(&mut output, 4, value.advanced.map(u64::from));
    if let Some(angle) = value.angle {
        let mut nested = Vec::with_capacity(5);
        put_float(&mut nested, 2, Some(angle));
        put_message(&mut output, 5, &nested);
    }
    output
}

fn encode_image(value: &ImageFill) -> Vec<u8> {
    let mut output = Vec::with_capacity(64);
    put_enum(&mut output, 2, value.technique);
    if let Some(tint) = &value.tint {
        put_message(&mut output, 3, &encode_color(tint));
    }
    if let Some([width, height]) = value.fill_size {
        let mut nested = Vec::with_capacity(10);
        put_float(&mut nested, 1, Some(width));
        put_float(&mut nested, 2, Some(height));
        put_message(&mut output, 4, &nested);
    }
    let mut reference = Vec::with_capacity(10);
    put_varint(&mut reference, 1, Some(value.data.unwrap_or(0)));
    put_message(&mut output, 6, &reference);
    put_varint(&mut output, 8, value.generic.map(u64::from));
    if let Some(color) = &value.reference_color {
        put_message(&mut output, 9, &encode_color(color));
    }
    output
}

/// Encode a `TSD.FillArchive` (an image fill always carries its data
/// reference, 0 for the placeholder; legacy image references are not
/// encodable).
#[must_use]
pub fn encode_fill(value: &Fill) -> Vec<u8> {
    let mut output = Vec::with_capacity(96);
    if let Some(color) = &value.color {
        put_message(&mut output, 1, &encode_color(color));
    }
    if let Some(gradient) = &value.gradient {
        put_message(&mut output, 2, &encode_gradient(gradient));
    }
    if let Some(image) = &value.image {
        put_message(&mut output, 3, &encode_image(image));
    }
    output
}
