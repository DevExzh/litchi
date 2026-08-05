//! Semantic Keynote slide transition settings and transactional editing.

use super::transition_wire::{patch_transition_settings_wire, transition_settings_from_wire};
use super::*;
use litchi_keynote::transition::{Acceleration, Direction, Effect, MosaicType, TextDelivery};

const SLIDE_ARCHIVE_MESSAGE_TYPE: u32 = 5;

/// Modern transition fields embedded in a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteTransitionSettings {
    pub animation_type: Option<String>,
    pub effect: Option<Effect>,
    pub duration: Option<f64>,
    pub direction: Option<Direction>,
    pub delay: Option<f64>,
    pub is_automatic: Option<bool>,
    /// Modern animation-level parameters, including byte-exact native color
    /// and timing-curve protobuf payloads.
    pub animation_parameters: KeynoteTransitionAnimationParameters,
    /// Native effect-specific parameters shared by transition effects.
    pub custom_parameters: KeynoteTransitionCustomParameters,
}

/// Lossless parameters stored inside a transition's modern animation archive.
///
/// Color and timing curves are kept as encoded native protobuf payloads. This
/// permits arbitrary current and future path-source variants to round-trip
/// without exposing private generated protobuf types in the public API.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct KeynoteTransitionAnimationParameters {
    pub color_payload: Option<Vec<u8>>,
    pub timing_curve_payloads: [Option<Vec<u8>>; 3],
    pub random_number_seed: Option<u32>,
    pub detail: Option<f64>,
    pub timing_curve_theme_names: [Option<String>; 3],
    pub writing_direction_is_rtl: Option<bool>,
}

/// Lossless native parameters shared by Keynote transition effects.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct KeynoteTransitionCustomParameters {
    pub twist: Option<f32>,
    pub mosaic_size: Option<u32>,
    pub mosaic_type: Option<MosaicType>,
    pub bounce: Option<bool>,
    pub magic_move_fade_unmatched_objects: Option<bool>,
    pub acceleration: Option<Acceleration>,
    pub text_delivery: Option<TextDelivery>,
    pub motion_blur: Option<bool>,
    pub travel_distance: Option<f32>,
}

impl KeynoteTransitionSettings {
    pub(super) fn from_native(attributes: &kn::TransitionAttributesArchive) -> Option<Self> {
        let animation = attributes.animation_attributes.as_ref()?;
        Some(Self {
            animation_type: animation.animation_type.clone(),
            effect: animation.effect.as_deref().map(Effect::from_identifier),
            duration: animation.duration,
            direction: animation.direction.map(Direction::from_native),
            delay: animation.delay,
            is_automatic: animation.is_automatic,
            animation_parameters: KeynoteTransitionAnimationParameters {
                color_payload: None,
                timing_curve_payloads: [None, None, None],
                random_number_seed: animation.random_number_seed,
                detail: animation.custom_detail,
                timing_curve_theme_names: [
                    animation.custom_effect_timing_curve_theme_name_1.clone(),
                    animation.custom_effect_timing_curve_theme_name_2.clone(),
                    animation.custom_effect_timing_curve_theme_name_3.clone(),
                ],
                writing_direction_is_rtl: animation.writing_direction_is_rtl,
            },
            custom_parameters: KeynoteTransitionCustomParameters {
                twist: attributes.custom_twist,
                mosaic_size: attributes.custom_mosaic_size,
                mosaic_type: attributes.custom_mosaic_type.map(MosaicType::from_native),
                bounce: attributes.custom_bounce,
                magic_move_fade_unmatched_objects: attributes
                    .custom_magic_move_fade_unmatched_objects,
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
        settings: KeynoteTransitionSettings,
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

fn validate_transition_settings(settings: &KeynoteTransitionSettings) -> Result<()> {
    for (name, value) in [
        ("transition duration", settings.duration),
        ("transition delay", settings.delay),
    ] {
        if value.is_some_and(|value| !value.is_finite() || value < 0.0) {
            return Err(Error::ParseError(format!(
                "Keynote {name} must be finite and non-negative"
            )));
        }
    }
    for (name, value) in [
        ("transition twist", settings.custom_parameters.twist),
        (
            "transition travel distance",
            settings.custom_parameters.travel_distance,
        ),
    ] {
        if value.is_some_and(|value| !value.is_finite()) {
            return Err(Error::ParseError(format!("Keynote {name} must be finite")));
        }
    }
    if settings
        .animation_parameters
        .detail
        .is_some_and(|value| !value.is_finite())
    {
        return Err(Error::ParseError(
            "Keynote transition detail must be finite".to_owned(),
        ));
    }
    if settings
        .effect
        .as_ref()
        .is_some_and(|effect| !effect.is_canonical())
    {
        return Err(Error::ParseError(
            "Keynote transition effect must use its named variant for known native identifiers"
                .to_owned(),
        ));
    }
    if settings
        .animation_type
        .as_deref()
        .into_iter()
        .chain(settings.effect.as_ref().map(Effect::identifier))
        .chain(
            settings
                .animation_parameters
                .timing_curve_theme_names
                .iter()
                .filter_map(Option::as_deref),
        )
        .any(|value| value.contains('\0'))
    {
        return Err(Error::ParseError(
            "Keynote transition strings cannot contain NUL".to_owned(),
        ));
    }
    if let Some(payload) = &settings.animation_parameters.color_payload {
        let color = tsp::Color::decode(payload.as_slice()).map_err(|error| {
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
        tsd::PathSourceArchive::decode(payload.as_slice()).map_err(|error| {
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
