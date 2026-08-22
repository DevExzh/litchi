//! Semantic Keynote slide-transition conversion and validation.

use super::*;
use litchi_iwa_protos::keynote_slide_transition_codec::{
    AnimationSnapshot, TransitionSettingsSnapshot,
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

    if let Some(payload) = settings.animation_parameters().color_payload() {
        let color = tsp::Color::decode(payload).map_err(|error| {
            Error::ParseError(format!("invalid Keynote transition color payload: {error}"))
        })?;
        for component in [
            color.r, color.g, color.b, color.a, color.c, color.m, color.y, color.k, color.w,
        ] {
            if component.is_some_and(|value| !value.is_finite()) {
                return Err(Error::ParseError(
                    "Keynote transition color components must be finite".to_owned(),
                ));
            }
        }
    }
    for payload in settings
        .animation_parameters()
        .timing_curve_payloads()
        .into_iter()
        .flatten()
    {
        tsd::PathSourceArchive::decode(payload).map_err(|error| {
            Error::ParseError(format!(
                "invalid Keynote transition timing-curve payload: {error}"
            ))
        })?;
    }
    Ok(())
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
}
