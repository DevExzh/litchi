//! Native protobuf wire translation for typed Keynote gradients.

use super::slide_background_color::{color_from_native, color_to_native};
use super::slide_background_gradient::{
    KeynoteGradient, KeynoteGradientAngle, KeynoteGradientKind, KeynoteGradientStop,
    validate_gradient,
};
use super::*;

const FILL_GRADIENT_FIELD: u32 = 2;
const GRADIENT_TYPE_FIELD: u32 = 1;
const GRADIENT_STOP_FIELD: u32 = 2;
const GRADIENT_OPACITY_FIELD: u32 = 3;
const GRADIENT_ADVANCED_FIELD: u32 = 4;
const GRADIENT_ANGLE_FIELD: u32 = 5;
const STOP_COLOR_FIELD: u32 = 1;
const STOP_POSITION_FIELD: u32 = 2;
const STOP_MIDPOINT_FIELD: u32 = 3;
const ANGLE_RADIANS_FIELD: u32 = 2;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_SPACE_FIELD: u32 = 12;

pub(super) fn gradient_from_fill(fill_payload: &[u8]) -> Result<Option<KeynoteGradient>> {
    if !has_exact_fields(fill_payload, &[FILL_GRADIENT_FIELD])? {
        return Ok(None);
    }
    let gradient_payload =
        required_length_delimited_payload(fill_payload, FILL_GRADIENT_FIELD, "Keynote gradient")?;
    let native = tsd::GradientArchive::decode(gradient_payload)?;
    let mut expected_gradient_fields = vec![
        GRADIENT_TYPE_FIELD,
        GRADIENT_OPACITY_FIELD,
        GRADIENT_ADVANCED_FIELD,
        GRADIENT_ANGLE_FIELD,
    ];
    expected_gradient_fields.extend(std::iter::repeat_n(GRADIENT_STOP_FIELD, native.stops.len()));
    if !has_exact_fields(gradient_payload, &expected_gradient_fields)? {
        return Ok(None);
    }
    let kind = match native
        .r#type
        .and_then(|value| tsd::gradient_archive::GradientType::try_from(value).ok())
    {
        Some(tsd::gradient_archive::GradientType::Linear) => KeynoteGradientKind::Linear,
        Some(tsd::gradient_archive::GradientType::Radial) => KeynoteGradientKind::Radial,
        None => return Ok(None),
    };
    if native.transformgradient.is_some() {
        return Ok(None);
    }
    let Some(opacity) = native.opacity else {
        return Ok(None);
    };
    let Some(is_advanced) = native.advanced_gradient else {
        return Ok(None);
    };
    let Some(angle_archive) = native.anglegradient else {
        return Ok(None);
    };
    let angle_payload = required_length_delimited_payload(
        gradient_payload,
        GRADIENT_ANGLE_FIELD,
        "Keynote gradient angle",
    )?;
    if !has_exact_fields(angle_payload, &[ANGLE_RADIANS_FIELD])? {
        return Ok(None);
    }
    let Some(angle_radians) = angle_archive.gradientangle else {
        return Ok(None);
    };
    let angle = match KeynoteGradientAngle::from_radians(angle_radians) {
        Ok(angle) => angle,
        Err(_) => return Ok(None),
    };
    let stop_payloads = repeated_length_delimited_payloads(gradient_payload, GRADIENT_STOP_FIELD)?;
    if stop_payloads.len() != native.stops.len() {
        return Ok(None);
    }
    let mut stops = Vec::with_capacity(native.stops.len());
    for (stop, payload) in native.stops.into_iter().zip(stop_payloads) {
        if !has_exact_fields(
            payload,
            &[STOP_COLOR_FIELD, STOP_POSITION_FIELD, STOP_MIDPOINT_FIELD],
        )? {
            return Ok(None);
        }
        let Some(native_color) = stop.color.as_ref() else {
            return Ok(None);
        };
        let color_payload =
            required_length_delimited_payload(payload, STOP_COLOR_FIELD, "Keynote gradient stop")?;
        if !has_exact_fields(
            color_payload,
            &[
                COLOR_MODEL_FIELD,
                COLOR_RED_FIELD,
                COLOR_GREEN_FIELD,
                COLOR_BLUE_FIELD,
                COLOR_ALPHA_FIELD,
                COLOR_SPACE_FIELD,
            ],
        )? {
            return Ok(None);
        }
        let Some(color) = color_from_native(native_color) else {
            return Ok(None);
        };
        let (Some(position), Some(midpoint)) = (stop.fraction, stop.inflection) else {
            return Ok(None);
        };
        stops.push(KeynoteGradientStop {
            color,
            position,
            midpoint,
        });
    }
    let gradient = KeynoteGradient {
        kind,
        stops,
        opacity,
        is_advanced,
        angle,
    };
    if validate_gradient(&gradient).is_err() {
        return Ok(None);
    }
    Ok(Some(gradient))
}

pub(super) fn gradient_to_fill(gradient: &KeynoteGradient) -> Vec<u8> {
    tsd::FillArchive {
        gradient: Some(tsd::GradientArchive {
            r#type: Some(match gradient.kind {
                KeynoteGradientKind::Linear => tsd::gradient_archive::GradientType::Linear as i32,
                KeynoteGradientKind::Radial => tsd::gradient_archive::GradientType::Radial as i32,
            }),
            stops: gradient
                .stops
                .iter()
                .map(|stop| tsd::gradient_archive::GradientStop {
                    color: Some(color_to_native(stop.color)),
                    fraction: Some(stop.position),
                    inflection: Some(stop.midpoint),
                })
                .collect(),
            opacity: Some(gradient.opacity),
            advanced_gradient: Some(gradient.is_advanced),
            anglegradient: Some(tsd::AngleGradientArchive {
                gradientangle: Some(gradient.angle.radians()),
            }),
            transformgradient: None,
        }),
        ..Default::default()
    }
    .encode_to_vec()
}

fn has_exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}
