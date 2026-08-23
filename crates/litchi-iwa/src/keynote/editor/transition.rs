//! Semantic Keynote slide-transition conversion and validation.

use super::*;
use litchi_iwa_protos::keynote_slide_transition_codec::{
    AnimationSnapshot, DecodeError as TransitionDecodeError,
    DecodeOptions as TransitionDecodeOptions, TransitionSettingsSnapshot, WireResourceLimit,
    validate_opaque_color, validate_opaque_path,
};
use litchi_keynote::transition::{
    Acceleration, AnimationParameters, CustomParameters, Direction, Effect, MosaicType,
    Settings as TransitionSettings, TextDelivery, TimingCurveSlot,
};

/// Convert a strict, borrowed transition projection into archive-free
/// semantic settings.
///
/// The projection preserves source presence and unknown enum discriminants.
/// Opaque color and timing-curve payloads are copied into the semantic value
/// byte-for-byte and validated by [`validate_transition_settings`]; the
/// generated transition archive is never used as a semantic authority.
pub(super) fn settings_from_projection(
    snapshot: &TransitionSettingsSnapshot<'_>,
) -> Result<Option<TransitionSettings>> {
    let Some(animation) = snapshot.animation else {
        return Ok(None);
    };

    let settings = settings_from_animation(snapshot, animation)?;
    validate_transition_settings(&settings)?;
    Ok(Some(settings))
}

fn settings_from_animation(
    snapshot: &TransitionSettingsSnapshot<'_>,
    animation: AnimationSnapshot<'_>,
) -> Result<TransitionSettings> {
    let effect = animation
        .effect
        .map(Effect::from_identifier)
        .transpose()
        .map_err(|error| transition_semantic_error("effect", error))?;

    let mut animation_parameters = AnimationParameters::new();
    animation_parameters.set_random_number_seed(animation.random_number_seed);
    animation_parameters
        .set_detail(animation.custom_detail)
        .map_err(|error| transition_semantic_error("detail", error))?;
    animation_parameters.set_writing_direction_is_rtl(animation.writing_direction_is_rtl);
    for (slot, theme_name) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_theme_name_1,
        animation.custom_effect_timing_curve_theme_name_2,
        animation.custom_effect_timing_curve_theme_name_3,
    ]) {
        animation_parameters
            .set_timing_curve_theme_name(slot, theme_name)
            .map_err(|error| transition_semantic_error("timing-curve theme", error))?;
    }
    animation_parameters
        .set_color_payload(animation.color)
        .map_err(|error| transition_semantic_error("color", error))?;
    for (slot, payload) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_1,
        animation.custom_effect_timing_curve_2,
        animation.custom_effect_timing_curve_3,
    ]) {
        animation_parameters
            .set_timing_curve_payload(slot, payload)
            .map_err(|error| transition_semantic_error("timing curve", error))?;
    }

    let mut custom_parameters = CustomParameters::new();
    custom_parameters
        .set_twist(snapshot.custom_twist)
        .map_err(|error| transition_semantic_error("twist", error))?;
    custom_parameters.set_mosaic_size(snapshot.custom_mosaic_size);
    custom_parameters.set_mosaic_type(snapshot.custom_mosaic_type.map(MosaicType::from_native));
    custom_parameters.set_bounce(snapshot.custom_bounce);
    custom_parameters
        .set_magic_move_fade_unmatched_objects(snapshot.custom_magic_move_fade_unmatched_objects);
    custom_parameters.set_acceleration(snapshot.custom_timing_curve.map(Acceleration::from_native));
    custom_parameters.set_text_delivery(
        snapshot
            .custom_text_delivery_type
            .map(TextDelivery::from_native),
    );
    custom_parameters.set_motion_blur(snapshot.custom_motion_blur);
    custom_parameters
        .set_travel_distance(snapshot.custom_travel_distance)
        .map_err(|error| transition_semantic_error("travel distance", error))?;

    let mut settings = TransitionSettings::new();
    settings
        .set_animation_type(animation.animation_type)
        .map_err(|error| transition_semantic_error("animation type", error))?;
    settings
        .set_effect(effect)
        .map_err(|error| transition_semantic_error("effect", error))?;
    settings
        .set_duration(animation.duration)
        .map_err(|error| transition_semantic_error("duration", error))?;
    settings.set_direction(animation.direction.map(Direction::from_native));
    settings
        .set_delay(animation.delay)
        .map_err(|error| transition_semantic_error("delay", error))?;
    settings.set_is_automatic(animation.is_automatic);
    settings
        .set_animation_parameters(animation_parameters)
        .map_err(|error| transition_semantic_error("animation parameters", error))?;
    settings
        .set_custom_parameters(custom_parameters)
        .map_err(|error| transition_semantic_error("custom parameters", error))?;
    Ok(settings)
}

fn transition_semantic_error(context: &str, error: impl std::fmt::Display) -> Error {
    Error::ParseError(format!("invalid Keynote transition {context}: {error}"))
}

pub(super) fn validate_transition_settings(settings: &TransitionSettings) -> Result<()> {
    settings.validate().map_err(|error| {
        Error::ParseError(format!("invalid Keynote transition settings: {error}"))
    })?;
    let opaque_options = opaque_transition_decode_options()?;
    let parameters = settings.animation_parameters();
    if let Some(payload) = parameters.color_payload() {
        validate_opaque_color(payload, opaque_options)
            .map_err(map_opaque_transition_codec_error)?;
    }
    for payload in parameters.timing_curve_payloads().into_iter().flatten() {
        validate_opaque_path(payload, opaque_options).map_err(map_opaque_transition_codec_error)?;
    }
    Ok(())
}

fn map_opaque_transition_codec_error(error: TransitionDecodeError) -> Error {
    const PREFIX: &str = "invalid Keynote transition opaque payload";
    let limit = match error.wire_resource_limit() {
        Some(WireResourceLimit::Bytes { observed, maximum }) => {
            Some(("wire bytes", observed as u64, maximum as u64))
        },
        Some(WireResourceLimit::Nesting { observed, maximum }) => {
            Some(("wire nesting", u64::from(observed), u64::from(maximum)))
        },
        Some(_) => None,
        None => error
            .field_limit_values()
            .map(|(observed, maximum)| ("wire fields", observed as u64, maximum as u64))
            .or_else(|| {
                error
                    .work_limit_values()
                    .map(|(observed, maximum)| ("wire work", observed as u64, maximum as u64))
            }),
    };
    if let Some((kind, observed, maximum)) = limit {
        return Error::ParseError(format!(
            "{PREFIX}: Keynote slide-transition {kind} limit exceeded: observed {observed}, maximum {maximum}"
        ));
    }
    Error::ParseError(format!(
        "{PREFIX}: the requested Keynote slide transition contains an invalid opaque payload"
    ))
}

fn opaque_transition_decode_options() -> Result<TransitionDecodeOptions> {
    let limits = WireLimits::default();
    let recursion_limit = u32::try_from(limits.max_nesting()).map_err(|_error| {
        Error::ParseError("Keynote transition nesting limit is out of range".to_owned())
    })?;
    // Opaque validation historically consumed the wire input/field/nesting
    // profile without imposing the rewrite-work ceiling. Bound the codec's
    // two strict scans by that same aggregate input budget.
    let validation_work = limits.max_input_bytes().saturating_mul(2);
    Ok(
        TransitionDecodeOptions::new(limits.max_input_bytes(), recursion_limit)
            .with_resource_limits(limits.max_fields(), validation_work),
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn transition_accelerations_map_native_values_losslessly() {
        for (raw, acceleration) in [
            (1, Acceleration::Linear),
            (2, Acceleration::EaseIn),
            (3, Acceleration::EaseOut),
            (4, Acceleration::EaseInOut),
            (5, Acceleration::Custom),
            (19, Acceleration::from_native(19)),
            (-1, Acceleration::from_native(-1)),
        ] {
            assert_eq!(Acceleration::from_native(raw), acceleration);
            assert_eq!(acceleration.native_value(), raw);
        }
    }

    #[test]
    fn transition_text_delivery_maps_native_values_losslessly() {
        for (raw, delivery) in [
            (1, TextDelivery::ByObject),
            (2, TextDelivery::ByWord),
            (3, TextDelivery::ByCharacter),
            (4, TextDelivery::ByLine),
            (19, TextDelivery::from_native(19)),
            (-1, TextDelivery::from_native(-1)),
        ] {
            assert_eq!(TextDelivery::from_native(raw), delivery);
            assert_eq!(delivery.native_value(), raw);
        }
    }

    #[test]
    fn effect_specific_discriminators_round_trip() {
        assert_eq!(Direction::from_native(42).native_value(), 42);
        assert_eq!(MosaicType::from_native(7).native_value(), 7);
    }

    #[test]
    fn strict_projection_adapter_retains_opaque_payloads_and_unknown_values() {
        let color = tsp::Color {
            model: tsp::color::ColorModel::White as i32,
            w: Some(0.5),
            ..Default::default()
        }
        .encode_to_vec();
        let timing_curve = tsd::PathSourceArchive {
            horizontal_flip: Some(true),
            ..Default::default()
        }
        .encode_to_vec();
        let animation = AnimationSnapshot {
            animation_type: Some("Transition"),
            effect: Some("com.example.future-transition"),
            duration: Some(1.25),
            direction: Some(42),
            delay: Some(0.5),
            is_automatic: Some(false),
            color: Some(color.as_slice()),
            custom_effect_timing_curve_1: Some(timing_curve.as_slice()),
            custom_effect_timing_curve_2: None,
            custom_effect_timing_curve_3: None,
            random_number_seed: Some(7),
            custom_detail: Some(2.0),
            custom_effect_timing_curve_theme_name_1: Some("future-theme"),
            custom_effect_timing_curve_theme_name_2: None,
            custom_effect_timing_curve_theme_name_3: None,
            writing_direction_is_rtl: Some(true),
        };
        let snapshot = TransitionSettingsSnapshot {
            has_legacy_database_fields: false,
            animation: Some(animation),
            custom_twist: Some(2.5),
            custom_mosaic_size: Some(9),
            custom_mosaic_type: Some(11),
            custom_bounce: Some(true),
            custom_magic_move_fade_unmatched_objects: Some(false),
            custom_timing_curve: Some(19),
            custom_text_delivery_type: Some(-1),
            custom_motion_blur: Some(true),
            custom_travel_distance: Some(88.0),
        };

        let settings = settings_from_projection(&snapshot)
            .expect("strict transition projection should map")
            .expect("modern animation should produce semantic settings");
        assert_eq!(
            settings.effect().map(Effect::identifier),
            Some("com.example.future-transition")
        );
        assert_eq!(settings.direction().map(Direction::native_value), Some(42));
        assert_eq!(
            settings.animation_parameters().color_payload(),
            Some(color.as_slice())
        );
        assert_eq!(
            settings
                .animation_parameters()
                .timing_curve_payload(TimingCurveSlot::First),
            Some(timing_curve.as_slice())
        );
        assert_eq!(
            settings
                .custom_parameters()
                .mosaic_type()
                .map(MosaicType::native_value),
            Some(11)
        );
        assert_eq!(
            settings
                .custom_parameters()
                .acceleration()
                .map(Acceleration::native_value),
            Some(19)
        );
        assert_eq!(
            settings
                .custom_parameters()
                .text_delivery()
                .map(TextDelivery::native_value),
            Some(-1)
        );
    }

    #[test]
    fn strict_projection_adapter_rejects_malformed_opaque_payloads() {
        let duplicate_color = [0x08, 0x01, 0x08, 0x01];
        let color_animation = AnimationSnapshot {
            animation_type: Some("Transition"),
            effect: Some("none"),
            duration: Some(1.0),
            direction: None,
            delay: None,
            is_automatic: None,
            color: Some(&duplicate_color),
            custom_effect_timing_curve_1: None,
            custom_effect_timing_curve_2: None,
            custom_effect_timing_curve_3: None,
            random_number_seed: None,
            custom_detail: None,
            custom_effect_timing_curve_theme_name_1: None,
            custom_effect_timing_curve_theme_name_2: None,
            custom_effect_timing_curve_theme_name_3: None,
            writing_direction_is_rtl: None,
        };
        let color_snapshot = TransitionSettingsSnapshot {
            has_legacy_database_fields: false,
            animation: Some(color_animation),
            custom_twist: None,
            custom_mosaic_size: None,
            custom_mosaic_type: None,
            custom_bounce: None,
            custom_magic_move_fade_unmatched_objects: None,
            custom_timing_curve: None,
            custom_text_delivery_type: None,
            custom_motion_blur: None,
            custom_travel_distance: None,
        };
        assert!(settings_from_projection(&color_snapshot).is_err());

        let invalid_path = [0x4a, 0x01, 0xff];
        for slot in 0..3 {
            let mut payloads = [None; 3];
            payloads[slot] = Some(invalid_path.as_slice());
            let path_animation = AnimationSnapshot {
                color: None,
                custom_effect_timing_curve_1: payloads[0],
                custom_effect_timing_curve_2: payloads[1],
                custom_effect_timing_curve_3: payloads[2],
                ..color_animation
            };
            let path_snapshot = TransitionSettingsSnapshot {
                animation: Some(path_animation),
                ..color_snapshot
            };
            assert!(settings_from_projection(&path_snapshot).is_err());
        }
    }

    #[test]
    fn opaque_codec_errors_keep_adapter_redaction_and_limit_details() {
        fn mapped_message(error: TransitionDecodeError) -> String {
            let Error::ParseError(message) = map_opaque_transition_codec_error(error) else {
                panic!("opaque codec failure must remain a parse error");
            };
            message
        }

        const REDACTED: &str = "invalid Keynote transition opaque payload: the requested Keynote slide transition contains an invalid opaque payload";
        let malformed = validate_opaque_color(
            &[0x08, 0x01, 0x08, 0x01],
            TransitionDecodeOptions::new(4, 8),
        )
        .expect_err("duplicate color model must fail");
        assert_eq!(mapped_message(malformed), REDACTED);
        let missing = validate_opaque_color(&[], TransitionDecodeOptions::new(1, 8))
            .expect_err("missing color model must fail");
        assert_eq!(mapped_message(missing), REDACTED);

        let oversized = validate_opaque_color(&[0x08, 0x01], TransitionDecodeOptions::new(1, 8))
            .expect_err("oversized color must fail");
        assert_eq!(
            mapped_message(oversized),
            "invalid Keynote transition opaque payload: Keynote slide-transition wire bytes limit exceeded: observed 2, maximum 1"
        );

        let too_many_fields = validate_opaque_color(
            &[0x08, 0x01],
            TransitionDecodeOptions::new(2, 8).with_resource_limits(0, 4),
        )
        .expect_err("zero field ceiling must fail");
        assert_eq!(
            mapped_message(too_many_fields),
            "invalid Keynote transition opaque payload: Keynote slide-transition wire fields limit exceeded: observed 1, maximum 0"
        );

        let too_much_work = validate_opaque_color(
            &[0x08, 0x01],
            TransitionDecodeOptions::new(2, 8).with_resource_limits(2, 3),
        )
        .expect_err("work ceiling below the strict scan cost must fail");
        assert_eq!(
            mapped_message(too_much_work),
            "invalid Keynote transition opaque payload: Keynote slide-transition wire work limit exceeded: observed 4, maximum 3"
        );

        let too_deep = validate_opaque_path(
            &[0x1a, 0x02, 0x12, 0x00],
            TransitionDecodeOptions::new(16, 1),
        )
        .expect_err("path nesting above the ceiling must fail");
        assert_eq!(
            mapped_message(too_deep),
            "invalid Keynote transition opaque payload: Keynote slide-transition wire nesting limit exceeded: observed 2, maximum 1"
        );
    }
}
