//! Semantic Keynote slide-transition conversion and validation.

use super::*;
use litchi_keynote::transition::{
    Acceleration, AnimationParameters, CustomParameters, Direction, Effect, MosaicType,
    Settings as TransitionSettings, TextDelivery, TimingCurveSlot,
};

pub(super) fn settings_from_native(
    attributes: &kn::TransitionAttributesArchive,
) -> Result<Option<TransitionSettings>> {
    let Some(animation) = attributes.animation_attributes.as_ref() else {
        return Ok(None);
    };

    let effect = animation
        .effect
        .as_deref()
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
        animation.custom_effect_timing_curve_theme_name_1.as_deref(),
        animation.custom_effect_timing_curve_theme_name_2.as_deref(),
        animation.custom_effect_timing_curve_theme_name_3.as_deref(),
    ]) {
        animation_parameters
            .set_timing_curve_theme_name(slot, theme_name)
            .map_err(|error| transition_semantic_error("timing-curve theme", error))?;
    }

    let mut custom_parameters = CustomParameters::new();
    custom_parameters
        .set_twist(attributes.custom_twist)
        .map_err(|error| transition_semantic_error("twist", error))?;
    custom_parameters.set_mosaic_size(attributes.custom_mosaic_size);
    custom_parameters.set_mosaic_type(attributes.custom_mosaic_type.map(MosaicType::from_native));
    custom_parameters.set_bounce(attributes.custom_bounce);
    custom_parameters
        .set_magic_move_fade_unmatched_objects(attributes.custom_magic_move_fade_unmatched_objects);
    custom_parameters.set_acceleration(
        attributes
            .custom_timing_curve
            .map(Acceleration::from_native),
    );
    custom_parameters.set_text_delivery(
        attributes
            .custom_text_delivery_type
            .map(TextDelivery::from_native),
    );
    custom_parameters.set_motion_blur(attributes.custom_motion_blur);
    custom_parameters
        .set_travel_distance(attributes.custom_travel_distance)
        .map_err(|error| transition_semantic_error("travel distance", error))?;

    let mut settings = TransitionSettings::new();
    settings
        .set_animation_type(animation.animation_type.as_deref())
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
    Ok(Some(settings))
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
}
