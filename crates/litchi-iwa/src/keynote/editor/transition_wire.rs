//! Lossless protobuf wire handling for Keynote slide transitions.

use super::transition::{settings_from_native, validate_transition_settings};
use super::*;
use litchi_keynote::transition::{Effect, Settings as TransitionSettings, TimingCurveSlot};

const SLIDE_TRANSITION_FIELD: u32 = 4;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;
const ANIMATION_ATTRIBUTES_FIELD: u32 = 8;

const ANIMATION_TYPE_FIELD: u32 = 1;
const EFFECT_FIELD: u32 = 2;
const DURATION_FIELD: u32 = 3;
const DIRECTION_FIELD: u32 = 4;
const DELAY_FIELD: u32 = 5;
const IS_AUTOMATIC_FIELD: u32 = 6;
const COLOR_FIELD: u32 = 7;
const TIMING_CURVE_FIELDS: [u32; 3] = [8, 9, 10];
const RANDOM_NUMBER_SEED_FIELD: u32 = 11;
const DETAIL_FIELD: u32 = 12;
const TIMING_CURVE_THEME_NAME_FIELDS: [u32; 3] = [13, 14, 15];
const WRITING_DIRECTION_IS_RTL_FIELD: u32 = 16;

const CUSTOM_TWIST_FIELD: u32 = 9;
const CUSTOM_MOSAIC_SIZE_FIELD: u32 = 10;
const CUSTOM_MOSAIC_TYPE_FIELD: u32 = 11;
const CUSTOM_BOUNCE_FIELD: u32 = 12;
const CUSTOM_MAGIC_MOVE_FADE_FIELD: u32 = 13;
const CUSTOM_ACCELERATION_FIELD: u32 = 15;
const CUSTOM_TEXT_DELIVERY_FIELD: u32 = 16;
const CUSTOM_MOTION_BLUR_FIELD: u32 = 17;
const CUSTOM_TRAVEL_DISTANCE_FIELD: u32 = 18;

const fn animation_path(field: u32) -> [u32; 4] {
    [
        SLIDE_TRANSITION_FIELD,
        TRANSITION_ATTRIBUTES_FIELD,
        ANIMATION_ATTRIBUTES_FIELD,
        field,
    ]
}

const fn custom_path(field: u32) -> [u32; 3] {
    [SLIDE_TRANSITION_FIELD, TRANSITION_ATTRIBUTES_FIELD, field]
}

pub(super) fn transition_settings_from_wire(
    original: &[u8],
    attributes: &kn::TransitionAttributesArchive,
) -> Result<Option<TransitionSettings>> {
    let Some(mut settings) = settings_from_native(attributes)? else {
        return Ok(None);
    };
    let transition = required_length_delimited_payload(
        original,
        SLIDE_TRANSITION_FIELD,
        "Keynote slide transition",
    )?;
    let attributes_wire = required_length_delimited_payload(
        transition,
        TRANSITION_ATTRIBUTES_FIELD,
        "Keynote transition attributes",
    )?;
    let animation = required_length_delimited_payload(
        attributes_wire,
        ANIMATION_ATTRIBUTES_FIELD,
        "Keynote modern transition attributes",
    )?;
    let mut animation_parameters = settings.animation_parameters().clone();
    animation_parameters
        .set_color_payload(optional_length_delimited_payload(animation, COLOR_FIELD)?)
        .map_err(|error| transition_parameter_error("color", error))?;
    for (index, field_number) in TIMING_CURVE_FIELDS.into_iter().enumerate() {
        animation_parameters
            .set_timing_curve_payload(
                TimingCurveSlot::ALL[index],
                optional_length_delimited_payload(animation, field_number)?,
            )
            .map_err(|error| transition_parameter_error("timing curve", error))?;
    }
    settings
        .set_animation_parameters(animation_parameters)
        .map_err(|error| transition_parameter_error("animation parameters", error))?;
    validate_transition_settings(&settings)?;
    Ok(Some(settings))
}

fn transition_parameter_error(context: &str, error: impl std::fmt::Display) -> Error {
    Error::ParseError(format!("invalid Keynote transition {context}: {error}"))
}

pub(super) fn patch_transition_settings_wire(
    original: &[u8],
    attributes: &kn::TransitionAttributesArchive,
    settings: &TransitionSettings,
) -> Result<Vec<u8>> {
    let animation = attributes.animation_attributes.as_ref().ok_or_else(|| {
        Error::InvalidFormat("Keynote transition has no modern attributes".to_owned())
    })?;
    let animation_parameters = settings.animation_parameters();
    let custom_parameters = settings.custom_parameters();
    let mut data = patch_nested_length_delimited_field(
        original,
        &animation_path(ANIMATION_TYPE_FIELD),
        animation.animation_type.is_some(),
        settings.animation_type().map(str::as_bytes),
    )?;
    data = patch_nested_length_delimited_field(
        &data,
        &animation_path(EFFECT_FIELD),
        animation.effect.is_some(),
        settings.effect().map(Effect::identifier).map(str::as_bytes),
    )?;
    for (field_number, current, replacement) in [
        (DURATION_FIELD, animation.duration, settings.duration()),
        (DELAY_FIELD, animation.delay, settings.delay()),
    ] {
        data = patch_nested_fixed64_field(
            &data,
            &animation_path(field_number),
            current.is_some(),
            replacement.map(f64::to_bits),
        )?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(DIRECTION_FIELD),
        animation.direction.is_some(),
        settings
            .direction()
            .map(|direction| u64::from(direction.native_value())),
    )?;
    data = patch_nested_varint_field(
        &data,
        &animation_path(IS_AUTOMATIC_FIELD),
        animation.is_automatic.is_some(),
        settings.is_automatic().map(u64::from),
    )?;
    data = patch_nested_length_delimited_field(
        &data,
        &animation_path(COLOR_FIELD),
        animation.color.is_some(),
        animation_parameters.color_payload(),
    )?;
    for (index, field_number, current) in [
        (
            0,
            TIMING_CURVE_FIELDS[0],
            animation.custom_effect_timing_curve_1.is_some(),
        ),
        (
            1,
            TIMING_CURVE_FIELDS[1],
            animation.custom_effect_timing_curve_2.is_some(),
        ),
        (
            2,
            TIMING_CURVE_FIELDS[2],
            animation.custom_effect_timing_curve_3.is_some(),
        ),
    ] {
        data = patch_nested_length_delimited_field(
            &data,
            &animation_path(field_number),
            current,
            animation_parameters.timing_curve_payload(TimingCurveSlot::ALL[index]),
        )?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(RANDOM_NUMBER_SEED_FIELD),
        animation.random_number_seed.is_some(),
        animation_parameters.random_number_seed().map(u64::from),
    )?;
    data = patch_nested_fixed64_field(
        &data,
        &animation_path(DETAIL_FIELD),
        animation.custom_detail.is_some(),
        animation_parameters.detail().map(f64::to_bits),
    )?;
    for (index, field_number, current) in [
        (
            0,
            TIMING_CURVE_THEME_NAME_FIELDS[0],
            animation.custom_effect_timing_curve_theme_name_1.is_some(),
        ),
        (
            1,
            TIMING_CURVE_THEME_NAME_FIELDS[1],
            animation.custom_effect_timing_curve_theme_name_2.is_some(),
        ),
        (
            2,
            TIMING_CURVE_THEME_NAME_FIELDS[2],
            animation.custom_effect_timing_curve_theme_name_3.is_some(),
        ),
    ] {
        data = patch_nested_length_delimited_field(
            &data,
            &animation_path(field_number),
            current,
            animation_parameters
                .timing_curve_theme_name(TimingCurveSlot::ALL[index])
                .map(str::as_bytes),
        )?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(WRITING_DIRECTION_IS_RTL_FIELD),
        animation.writing_direction_is_rtl.is_some(),
        animation_parameters
            .writing_direction_is_rtl()
            .map(u64::from),
    )?;
    for (field_number, current, replacement) in [
        (
            CUSTOM_TWIST_FIELD,
            attributes.custom_twist,
            custom_parameters.twist(),
        ),
        (
            CUSTOM_TRAVEL_DISTANCE_FIELD,
            attributes.custom_travel_distance,
            custom_parameters.travel_distance(),
        ),
    ] {
        data = patch_nested_fixed32_field(
            &data,
            &custom_path(field_number),
            current.is_some(),
            replacement.map(f32::to_bits),
        )?;
    }
    let mosaic_type = custom_parameters
        .mosaic_type()
        .map(|value| u64::from(value.native_value()));
    let acceleration = custom_parameters
        .acceleration()
        .map(|value| i64::from(value.native_value()) as u64);
    let text_delivery = custom_parameters
        .text_delivery()
        .map(|value| i64::from(value.native_value()) as u64);
    for (field_number, current, replacement) in [
        (
            CUSTOM_MOSAIC_SIZE_FIELD,
            attributes.custom_mosaic_size.map(u64::from),
            custom_parameters.mosaic_size().map(u64::from),
        ),
        (
            CUSTOM_MOSAIC_TYPE_FIELD,
            attributes.custom_mosaic_type.map(u64::from),
            mosaic_type,
        ),
        (
            CUSTOM_BOUNCE_FIELD,
            attributes.custom_bounce.map(u64::from),
            custom_parameters.bounce().map(u64::from),
        ),
        (
            CUSTOM_MAGIC_MOVE_FADE_FIELD,
            attributes
                .custom_magic_move_fade_unmatched_objects
                .map(u64::from),
            custom_parameters
                .magic_move_fade_unmatched_objects()
                .map(u64::from),
        ),
        (
            CUSTOM_ACCELERATION_FIELD,
            attributes
                .custom_timing_curve
                .map(|value| i64::from(value) as u64),
            acceleration,
        ),
        (
            CUSTOM_TEXT_DELIVERY_FIELD,
            attributes
                .custom_text_delivery_type
                .map(|value| i64::from(value) as u64),
            text_delivery,
        ),
        (
            CUSTOM_MOTION_BLUR_FIELD,
            attributes.custom_motion_blur.map(u64::from),
            custom_parameters.motion_blur().map(u64::from),
        ),
    ] {
        data = patch_nested_varint_field(
            &data,
            &custom_path(field_number),
            current.is_some(),
            replacement,
        )?;
    }
    Ok(data)
}

pub(super) fn validate_transition_wire(
    original: &[u8],
    attributes: &kn::TransitionAttributesArchive,
) -> Result<()> {
    let settings = transition_settings_from_wire(original, attributes)?.ok_or_else(|| {
        Error::InvalidFormat("Keynote transition has no modern attributes".to_owned())
    })?;
    let verified = patch_transition_settings_wire(original, attributes, &settings)?;
    if verified != original {
        return Err(Error::InvalidFormat(
            "Keynote transition no-op wire validation changed its payload".to_owned(),
        ));
    }
    Ok(())
}
