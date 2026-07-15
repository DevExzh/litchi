//! Semantic Keynote slide transition settings and transactional editing.

use super::transition_wire::{patch_transition_settings_wire, transition_settings_from_wire};
use super::*;

const SLIDE_ARCHIVE_MESSAGE_TYPE: u32 = 5;

const LINEAR_ACCELERATION: i32 = 1;
const EASE_IN_ACCELERATION: i32 = 2;
const EASE_OUT_ACCELERATION: i32 = 3;
const EASE_IN_OUT_ACCELERATION: i32 = 4;
const CUSTOM_ACCELERATION: i32 = 5;

const BY_OBJECT_DELIVERY: i32 = 1;
const BY_WORD_DELIVERY: i32 = 2;
const BY_CHARACTER_DELIVERY: i32 = 3;
const BY_LINE_DELIVERY: i32 = 4;

/// Effect-specific direction discriminator stored by Keynote.
///
/// Direction values are interpreted by the selected effect, so this newtype
/// prevents unrelated integers from being passed accidentally while retaining
/// values introduced by future Keynote versions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct KeynoteTransitionDirection(u32);

impl KeynoteTransitionDirection {
    /// Wrap a native effect-specific direction value.
    pub const fn from_raw(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific direction value.
    pub const fn as_raw(self) -> u32 {
        self.0
    }
}

/// Effect-specific mosaic layout discriminator stored by Keynote.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct KeynoteTransitionMosaicType(u32);

impl KeynoteTransitionMosaicType {
    /// Wrap a native effect-specific mosaic layout value.
    pub const fn from_raw(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific mosaic layout value.
    pub const fn as_raw(self) -> u32 {
        self.0
    }
}

/// Acceleration curve used by transitions such as Magic Move.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteTransitionAcceleration {
    /// Constant speed; displayed as “None” in Keynote.
    Linear,
    EaseIn,
    EaseOut,
    EaseInOut,
    Custom,
    /// A value introduced by a newer Keynote version.
    Unknown(i32),
}

impl KeynoteTransitionAcceleration {
    /// Decode the native timing-curve enum without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            LINEAR_ACCELERATION => Self::Linear,
            EASE_IN_ACCELERATION => Self::EaseIn,
            EASE_OUT_ACCELERATION => Self::EaseOut,
            EASE_IN_OUT_ACCELERATION => Self::EaseInOut,
            CUSTOM_ACCELERATION => Self::Custom,
            value => Self::Unknown(value),
        }
    }

    /// Return the native transition timing-curve enum value.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::Linear => LINEAR_ACCELERATION,
            Self::EaseIn => EASE_IN_ACCELERATION,
            Self::EaseOut => EASE_OUT_ACCELERATION,
            Self::EaseInOut => EASE_IN_OUT_ACCELERATION,
            Self::Custom => CUSTOM_ACCELERATION,
            Self::Unknown(value) => value,
        }
    }

    const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(
                LINEAR_ACCELERATION
                    | EASE_IN_ACCELERATION
                    | EASE_OUT_ACCELERATION
                    | EASE_IN_OUT_ACCELERATION
                    | CUSTOM_ACCELERATION
            )
        )
    }
}

/// How matching text is delivered during a Keynote transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum KeynoteTransitionTextDelivery {
    ByObject,
    ByWord,
    ByCharacter,
    ByLine,
    /// A value introduced by a newer Keynote version.
    Unknown(i32),
}

impl KeynoteTransitionTextDelivery {
    /// Decode the native text-delivery enum without discarding unknown values.
    pub const fn from_raw(value: i32) -> Self {
        match value {
            BY_OBJECT_DELIVERY => Self::ByObject,
            BY_WORD_DELIVERY => Self::ByWord,
            BY_CHARACTER_DELIVERY => Self::ByCharacter,
            BY_LINE_DELIVERY => Self::ByLine,
            value => Self::Unknown(value),
        }
    }

    /// Return the native transition text-delivery enum value.
    pub const fn as_raw(self) -> i32 {
        match self {
            Self::ByObject => BY_OBJECT_DELIVERY,
            Self::ByWord => BY_WORD_DELIVERY,
            Self::ByCharacter => BY_CHARACTER_DELIVERY,
            Self::ByLine => BY_LINE_DELIVERY,
            Self::Unknown(value) => value,
        }
    }

    const fn is_canonical(self) -> bool {
        !matches!(
            self,
            Self::Unknown(
                BY_OBJECT_DELIVERY | BY_WORD_DELIVERY | BY_CHARACTER_DELIVERY | BY_LINE_DELIVERY
            )
        )
    }
}

/// Modern transition fields embedded in a Keynote slide.
#[derive(Debug, Clone, PartialEq)]
pub struct KeynoteTransitionSettings {
    pub animation_type: Option<String>,
    pub effect: Option<KeynoteTransitionEffect>,
    pub duration: Option<f64>,
    pub direction: Option<KeynoteTransitionDirection>,
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
    pub mosaic_type: Option<KeynoteTransitionMosaicType>,
    pub bounce: Option<bool>,
    pub magic_move_fade_unmatched_objects: Option<bool>,
    pub acceleration: Option<KeynoteTransitionAcceleration>,
    pub text_delivery: Option<KeynoteTransitionTextDelivery>,
    pub motion_blur: Option<bool>,
    pub travel_distance: Option<f32>,
}

impl KeynoteTransitionSettings {
    pub(super) fn from_native(attributes: &kn::TransitionAttributesArchive) -> Option<Self> {
        let animation = attributes.animation_attributes.as_ref()?;
        Some(Self {
            animation_type: animation.animation_type.clone(),
            effect: animation
                .effect
                .as_deref()
                .map(KeynoteTransitionEffect::from_identifier),
            duration: animation.duration,
            direction: animation
                .direction
                .map(KeynoteTransitionDirection::from_raw),
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
                mosaic_type: attributes
                    .custom_mosaic_type
                    .map(KeynoteTransitionMosaicType::from_raw),
                bounce: attributes.custom_bounce,
                magic_move_fade_unmatched_objects: attributes
                    .custom_magic_move_fade_unmatched_objects,
                acceleration: attributes
                    .custom_timing_curve
                    .map(KeynoteTransitionAcceleration::from_raw),
                text_delivery: attributes
                    .custom_text_delivery_type
                    .map(KeynoteTransitionTextDelivery::from_raw),
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
    /// typed effect is [`KeynoteTransitionEffect::None`], so replacing that
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
        .custom_parameters
        .acceleration
        .is_some_and(|value| !value.is_canonical())
    {
        return Err(Error::ParseError(
            "Keynote transition acceleration must use its named variant for known native values"
                .to_owned(),
        ));
    }
    if settings
        .custom_parameters
        .text_delivery
        .is_some_and(|value| !value.is_canonical())
    {
        return Err(Error::ParseError(
            "Keynote transition text delivery must use its named variant for known native values"
                .to_owned(),
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
        .chain(
            settings
                .effect
                .as_ref()
                .map(KeynoteTransitionEffect::as_identifier),
        )
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
            (1, KeynoteTransitionAcceleration::Linear),
            (2, KeynoteTransitionAcceleration::EaseIn),
            (3, KeynoteTransitionAcceleration::EaseOut),
            (4, KeynoteTransitionAcceleration::EaseInOut),
            (5, KeynoteTransitionAcceleration::Custom),
            (19, KeynoteTransitionAcceleration::Unknown(19)),
            (-1, KeynoteTransitionAcceleration::Unknown(-1)),
        ] {
            assert_eq!(KeynoteTransitionAcceleration::from_raw(raw), acceleration);
            assert_eq!(acceleration.as_raw(), raw);
        }
    }

    #[test]
    fn transition_text_delivery_maps_native_values_losslessly() {
        for (raw, delivery) in [
            (1, KeynoteTransitionTextDelivery::ByObject),
            (2, KeynoteTransitionTextDelivery::ByWord),
            (3, KeynoteTransitionTextDelivery::ByCharacter),
            (4, KeynoteTransitionTextDelivery::ByLine),
            (19, KeynoteTransitionTextDelivery::Unknown(19)),
            (-1, KeynoteTransitionTextDelivery::Unknown(-1)),
        ] {
            assert_eq!(KeynoteTransitionTextDelivery::from_raw(raw), delivery);
            assert_eq!(delivery.as_raw(), raw);
        }
    }

    #[test]
    fn effect_specific_discriminators_round_trip() {
        assert_eq!(KeynoteTransitionDirection::from_raw(42).as_raw(), 42);
        assert_eq!(KeynoteTransitionMosaicType::from_raw(7).as_raw(), 7);
    }
}
