//! Semantic Keynote slide transition settings and transactional editing.

use super::transition_wire::{patch_transition_settings_wire, transition_settings_from_wire};
use super::*;
use litchi_keynote::transition::{
    Acceleration, AnimationParameters, CustomParameters, Direction, Effect, MosaicType,
    Settings as TransitionSettings, TextDelivery,
};

const SLIDE_ARCHIVE_MESSAGE_TYPE: u32 = 5;

pub(super) fn settings_from_native(
    attributes: &kn::TransitionAttributesArchive,
) -> Option<TransitionSettings> {
    let animation = attributes.animation_attributes.as_ref()?;
    Some(TransitionSettings {
        animation_type: animation.animation_type.clone().map(String::into_boxed_str),
        effect: animation.effect.as_deref().map(Effect::from_identifier),
        duration: animation.duration,
        direction: animation.direction.map(Direction::from_native),
        delay: animation.delay,
        is_automatic: animation.is_automatic,
        animation_parameters: AnimationParameters {
            color_payload: None,
            timing_curve_payloads: [None, None, None],
            random_number_seed: animation.random_number_seed,
            detail: animation.custom_detail,
            timing_curve_theme_names: [
                animation
                    .custom_effect_timing_curve_theme_name_1
                    .clone()
                    .map(String::into_boxed_str),
                animation
                    .custom_effect_timing_curve_theme_name_2
                    .clone()
                    .map(String::into_boxed_str),
                animation
                    .custom_effect_timing_curve_theme_name_3
                    .clone()
                    .map(String::into_boxed_str),
            ],
            writing_direction_is_rtl: animation.writing_direction_is_rtl,
        },
        custom_parameters: CustomParameters {
            twist: attributes.custom_twist,
            mosaic_size: attributes.custom_mosaic_size,
            mosaic_type: attributes.custom_mosaic_type.map(MosaicType::from_native),
            bounce: attributes.custom_bounce,
            magic_move_fade_unmatched_objects: attributes.custom_magic_move_fade_unmatched_objects,
            acceleration: attributes
                .custom_timing_curve
                .map(Acceleration::from_native),
            text_delivery: attributes
                .custom_text_delivery_type
                .map(TextDelivery::from_native),
            motion_blur: attributes.custom_motion_blur,
            travel_distance: attributes.custom_travel_distance,
        },
    })
}

impl KeynoteEditor {
    /// Create or replace a slide's modern transition fields transactionally.
    ///
    /// Current Keynote slides encode “no effect” as a modern transition whose
    /// typed effect is [`Effect::None`], so replacing that
    /// value creates a visible transition without fabricating a new archive.
    ///
    /// Legacy-only transitions are rejected so the editor never guesses which
    /// representation should take precedence.
    pub fn set_slide_transition(
        &mut self,
        slide_index: usize,
        settings: TransitionSettings,
    ) -> Result<()> {
        validate_transition_settings(&settings)?;
        let slides = self.slides()?;
        let slide = slides.get(slide_index).ok_or_else(|| {
            Error::ParseError(format!(
                "Keynote slide index {slide_index} is out of range for {} slides",
                slides.len()
            ))
        })?;
        if slide.transition.is_none() {
            return Err(Error::InvalidFormat(format!(
                "Keynote slide {slide_index} has no modern transition attributes"
            )));
        }
        let slide_id = slide.slide_id;
        let graph = ObjectGraph::read(self.text.package())?;
        let archive_name = graph.archive_name(slide_id)?.to_owned();
        let mut staged = self.text.package().clone();
        staged.update_archive(&archive_name, |archive| {
            let object = archive.object_mut(slide_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote slide object {slide_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    message.type_ == SLIDE_ARCHIVE_MESSAGE_TYPE
                        && kn::SlideArchive::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Keynote slide object {slide_id} has no SlideArchive payload"
                    ))
                })?;
            let original = object.messages[message_index].data.as_slice();
            let slide = kn::SlideArchive::decode(original)?;
            let attributes = &slide.transition.attributes;
            if attributes.animation_attributes.is_none() {
                return Err(Error::InvalidFormat(format!(
                    "Keynote slide {slide_index} has no modern transition attributes"
                )));
            }
            let data = patch_transition_settings_wire(original, attributes, &settings)?;
            let verified = kn::SlideArchive::decode(data.as_slice())?;
            let verified = transition_settings_from_wire(&data, &verified.transition.attributes)?;
            if verified.as_ref() != Some(&settings) {
                return Err(Error::InvalidFormat(
                    "Keynote transition wire patch failed validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: SLIDE_ARCHIVE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })?;
        let verified = Self::from_bytes(&staged.to_bytes()?)?;
        if verified
            .slides()?
            .get(slide_index)
            .and_then(|slide| slide.transition.as_ref())
            != Some(&settings)
        {
            return Err(Error::InvalidFormat(
                "Keynote transition update failed round-trip validation".to_owned(),
            ));
        }
        self.text = IWorkTextEditor::from_package(staged);
        Ok(())
    }
}

pub(super) fn validate_transition_settings(settings: &TransitionSettings) -> Result<()> {
    settings.validate().map_err(|error| {
        Error::ParseError(format!("invalid Keynote transition settings: {error}"))
    })?;

    if let Some(payload) = &settings.animation_parameters.color_payload {
        let color = tsp::Color::decode(payload.as_ref()).map_err(|error| {
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
        .animation_parameters
        .timing_curve_payloads
        .iter()
        .flatten()
    {
        tsd::PathSourceArchive::decode(payload.as_ref()).map_err(|error| {
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
