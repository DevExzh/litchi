//! Keynote transition vocabulary independent of archive and wire models.

use crate::{Error, Result};

const NONE_EFFECT: &str = "none";
const DISSOLVE_EFFECT: &str = "apple:dissolve";
const MAGIC_MOVE_EFFECT: &str = "apple:magic-move-implied-motion-path";
const LINEAR_ACCELERATION: i32 = 1;
const EASE_IN_ACCELERATION: i32 = 2;
const EASE_OUT_ACCELERATION: i32 = 3;
const EASE_IN_OUT_ACCELERATION: i32 = 4;
const CUSTOM_ACCELERATION: i32 = 5;
const BY_OBJECT_DELIVERY: i32 = 1;
const BY_WORD_DELIVERY: i32 = 2;
const BY_CHARACTER_DELIVERY: i32 = 3;
const BY_LINE_DELIVERY: i32 = 4;
const TIMING_CURVE_SLOT_COUNT: usize = 3;

/// A slide-transition effect understood by Keynote.
///
/// Named variants cover identifiers verified in native Keynote documents.
/// [`Effect::Unknown`] preserves identifiers introduced by future releases
/// without reducing the semantic API to an untyped string.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum Effect {
    /// Keynote's native “No Transition Effect” representation.
    None,
    /// A dissolve transition.
    Dissolve,
    /// Keynote's Magic Move transition.
    MagicMove,
    /// An identifier introduced by a newer Keynote release.
    Unknown(Box<str>),
}

impl Effect {
    /// Decode a native effect identifier without discarding unknown values.
    #[must_use]
    pub fn from_identifier(identifier: &str) -> Self {
        match identifier {
            NONE_EFFECT => Self::None,
            DISSOLVE_EFFECT => Self::Dissolve,
            MAGIC_MOVE_EFFECT => Self::MagicMove,
            other => Self::Unknown(other.into()),
        }
    }

    /// Construct an unknown effect identifier while rejecting non-canonical
    /// known identifiers and strings that cannot be written to native text
    /// fields.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalEffect`] when `identifier` names a known
    /// effect, or [`Error::NulString`] when it contains a NUL byte.
    pub fn unknown(identifier: impl Into<Box<str>>) -> Result<Self> {
        let effect = Self::Unknown(identifier.into());
        effect.validate()?;
        Ok(effect)
    }

    /// Return the native Keynote effect identifier.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::None => NONE_EFFECT,
            Self::Dissolve => DISSOLVE_EFFECT,
            Self::MagicMove => MAGIC_MOVE_EFFECT,
            Self::Unknown(identifier) => identifier,
        }
    }

    /// Return whether the value uses a named variant for a known identifier.
    #[must_use]
    pub fn is_canonical(&self) -> bool {
        !matches!(
            self,
            Self::Unknown(identifier)
                if matches!(
                    &**identifier,
                    NONE_EFFECT | DISSOLVE_EFFECT | MAGIC_MOVE_EFFECT
                )
        )
    }

    /// Validate the effect before it is published through a native adapter.
    ///
    /// Unknown identifiers are valid and remain lossless. The unknown form is
    /// rejected only when it shadows a named native identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NulString`] when the identifier contains a NUL byte or
    /// [`Error::NonCanonicalEffect`] when an unknown value shadows a named
    /// native effect.
    pub fn validate(&self) -> Result<()> {
        if self.identifier().contains('\0') {
            return Err(Error::NulString);
        }
        if !self.is_canonical() {
            return Err(Error::NonCanonicalEffect);
        }
        Ok(())
    }
}

/// Effect-specific direction discriminator stored by Keynote.
///
/// The native value is retained directly so unknown effect-specific values
/// remain lossless without increasing the value beyond four bytes.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Direction(u32);

impl Direction {
    /// Wrap a native effect-specific direction value.
    #[must_use]
    pub const fn from_native(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific direction value.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Effect-specific mosaic layout discriminator stored by Keynote.
///
/// The native value is retained directly so unknown effect-specific values
/// remain lossless without increasing the value beyond four bytes.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct MosaicType(u32);

impl MosaicType {
    /// Wrap a native effect-specific mosaic layout value.
    #[must_use]
    pub const fn from_native(value: u32) -> Self {
        Self(value)
    }

    /// Return the native effect-specific mosaic layout value.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Recognized transition acceleration curves.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AccelerationKind {
    /// Constant speed; displayed as “None” in Keynote.
    Linear,
    /// Ease into the transition.
    EaseIn,
    /// Ease out of the transition.
    EaseOut,
    /// Ease into and out of the transition.
    EaseInOut,
    /// A custom timing curve.
    Custom,
}

/// Acceleration curve used by transitions such as Magic Move.
///
/// Known curves have named associated constants, while future native values
/// remain lossless in the same four-byte representation.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Acceleration(i32);

impl Acceleration {
    /// Constant-speed transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Linear: Self = Self(LINEAR_ACCELERATION);
    /// Ease into the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseIn: Self = Self(EASE_IN_ACCELERATION);
    /// Ease out of the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseOut: Self = Self(EASE_OUT_ACCELERATION);
    /// Ease into and out of the transition.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseInOut: Self = Self(EASE_IN_OUT_ACCELERATION);
    /// Use a custom timing curve.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const Custom: Self = Self(CUSTOM_ACCELERATION);

    /// Decode a native timing-curve value without discarding unknown values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native transition timing-curve value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized timing-curve kind, if known.
    #[must_use]
    pub const fn kind(self) -> Option<AccelerationKind> {
        match self.0 {
            LINEAR_ACCELERATION => Some(AccelerationKind::Linear),
            EASE_IN_ACCELERATION => Some(AccelerationKind::EaseIn),
            EASE_OUT_ACCELERATION => Some(AccelerationKind::EaseOut),
            EASE_IN_OUT_ACCELERATION => Some(AccelerationKind::EaseInOut),
            CUSTOM_ACCELERATION => Some(AccelerationKind::Custom),
            _ => None,
        }
    }
}

impl std::fmt::Debug for Acceleration {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.kind() {
            Some(kind) => kind.fmt(formatter),
            None => formatter.debug_tuple("Unsupported").field(&self.0).finish(),
        }
    }
}

/// Recognized ways to deliver matching text during a transition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextDeliveryKind {
    /// Deliver by object.
    ByObject,
    /// Deliver by word.
    ByWord,
    /// Deliver by character.
    ByCharacter,
    /// Deliver by line.
    ByLine,
}

/// How matching text is delivered during a Keynote transition.
///
/// Known delivery modes have named associated constants, while future native
/// values remain lossless in the same four-byte representation.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextDelivery(i32);

impl TextDelivery {
    /// Deliver by object.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByObject: Self = Self(BY_OBJECT_DELIVERY);
    /// Deliver by word.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByWord: Self = Self(BY_WORD_DELIVERY);
    /// Deliver by character.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByCharacter: Self = Self(BY_CHARACTER_DELIVERY);
    /// Deliver by line.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const ByLine: Self = Self(BY_LINE_DELIVERY);

    /// Decode a native text-delivery value without discarding unknown values.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native transition text-delivery value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized delivery kind, if known.
    #[must_use]
    pub const fn kind(self) -> Option<TextDeliveryKind> {
        match self.0 {
            BY_OBJECT_DELIVERY => Some(TextDeliveryKind::ByObject),
            BY_WORD_DELIVERY => Some(TextDeliveryKind::ByWord),
            BY_CHARACTER_DELIVERY => Some(TextDeliveryKind::ByCharacter),
            BY_LINE_DELIVERY => Some(TextDeliveryKind::ByLine),
            _ => None,
        }
    }
}

impl std::fmt::Debug for TextDelivery {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self.kind() {
            Some(kind) => kind.fmt(formatter),
            None => formatter.debug_tuple("Unsupported").field(&self.0).finish(),
        }
    }
}

/// Lossless animation-level values attached to a slide transition.
///
/// The payload fields intentionally remain opaque. Their protobuf structure,
/// validation, and wire-preserving patching belong to the IWA adapter; this
/// crate only owns bounded native-free storage and semantic scalar checks.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct AnimationParameters {
    /// Exact native color payload, when one is present.
    pub color_payload: Option<Box<[u8]>>,
    /// Exact native timing-curve payloads for Keynote's three curve slots.
    ///
    /// This is a fixed-size array because the native schema has exactly three
    /// slots. Keeping the slots inline avoids an allocation and capacity
    /// overhead for the overwhelmingly common all-absent value.
    pub timing_curve_payloads: [Option<Box<[u8]>>; TIMING_CURVE_SLOT_COUNT],
    /// Native random seed used by effects with randomized motion.
    pub random_number_seed: Option<u32>,
    /// Native effect detail value.
    pub detail: Option<f64>,
    /// Native theme names for the three timing-curve slots.
    pub timing_curve_theme_names: [Option<Box<str>>; TIMING_CURVE_SLOT_COUNT],
    /// Whether transition text should use right-to-left writing direction.
    pub writing_direction_is_rtl: Option<bool>,
}

impl AnimationParameters {
    /// Validate semantic values without interpreting opaque native payloads.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidDetail`] for a non-finite detail value or
    /// [`Error::NulString`] for a theme name containing a NUL byte.
    pub fn validate(&self) -> Result<()> {
        if self.detail.is_some_and(|value| !value.is_finite()) {
            return Err(Error::InvalidDetail);
        }
        if self
            .timing_curve_theme_names
            .iter()
            .flatten()
            .any(|name| name.contains('\0'))
        {
            return Err(Error::NulString);
        }
        Ok(())
    }

    /// Borrow an optional color payload without copying it.
    #[must_use]
    pub fn color_payload(&self) -> Option<&[u8]> {
        self.color_payload.as_deref()
    }

    /// Borrow the exact timing-curve payload slots.
    #[must_use]
    pub const fn timing_curve_payloads(&self) -> &[Option<Box<[u8]>>; TIMING_CURVE_SLOT_COUNT] {
        &self.timing_curve_payloads
    }

    /// Borrow the exact timing-curve theme-name slots.
    #[must_use]
    pub const fn timing_curve_theme_names(&self) -> &[Option<Box<str>>; TIMING_CURVE_SLOT_COUNT] {
        &self.timing_curve_theme_names
    }
}

/// Lossless effect-specific values shared by Keynote transitions.
#[derive(Debug, Clone, Copy, Default, PartialEq)]
pub struct CustomParameters {
    /// Native twist amount, when supplied by the effect.
    pub twist: Option<f32>,
    /// Native mosaic size, when supplied by the effect.
    pub mosaic_size: Option<u32>,
    /// Native mosaic layout discriminator.
    pub mosaic_type: Option<MosaicType>,
    /// Whether a bounce effect is enabled.
    pub bounce: Option<bool>,
    /// Whether Magic Move fades unmatched objects.
    pub magic_move_fade_unmatched_objects: Option<bool>,
    /// Native timing-curve discriminator.
    pub acceleration: Option<Acceleration>,
    /// Native matching-text delivery discriminator.
    pub text_delivery: Option<TextDelivery>,
    /// Whether motion blur is enabled.
    pub motion_blur: Option<bool>,
    /// Native travel distance, when supplied by the effect.
    pub travel_distance: Option<f32>,
}

impl CustomParameters {
    /// Validate all custom floating-point values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidCustomFloat`] when a custom value is NaN or
    /// infinite. Unknown integer discriminators remain valid and lossless.
    pub fn validate(&self) -> Result<()> {
        if self
            .twist
            .into_iter()
            .chain(self.travel_distance)
            .any(|value| !value.is_finite())
        {
            return Err(Error::InvalidCustomFloat);
        }
        Ok(())
    }
}

/// Lossless, archive-free semantic settings for one Keynote slide transition.
///
/// The fields retain native presence and unknown values, while the owned
/// string and payload fields use exact-size boxed storage. Call [`Self::validate`]
/// before passing settings to an IWA adapter or publishing them from an edit.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Settings {
    /// Native animation name, when present.
    pub animation_type: Option<Box<str>>,
    /// Semantic transition effect.
    pub effect: Option<Effect>,
    /// Transition duration in seconds.
    pub duration: Option<f64>,
    /// Effect-specific direction discriminator.
    pub direction: Option<Direction>,
    /// Delay before the transition starts, in seconds.
    pub delay: Option<f64>,
    /// Whether the transition starts automatically.
    pub is_automatic: Option<bool>,
    /// Animation-level opaque payloads and semantic values.
    pub animation_parameters: AnimationParameters,
    /// Effect-specific scalar values.
    pub custom_parameters: CustomParameters,
}

impl Settings {
    /// Validate every semantic field without decoding opaque payloads.
    ///
    /// # Errors
    ///
    /// Returns a typed [`Error`] when a duration or delay is non-finite or
    /// negative, a custom value or detail is non-finite, a string contains a
    /// NUL byte, or an effect uses a non-canonical known identifier.
    pub fn validate(&self) -> Result<()> {
        if self
            .duration
            .is_some_and(|value| !value.is_finite() || value < 0.0)
        {
            return Err(Error::InvalidDuration);
        }
        if self
            .delay
            .is_some_and(|value| !value.is_finite() || value < 0.0)
        {
            return Err(Error::InvalidDelay);
        }
        if self
            .animation_type
            .as_deref()
            .is_some_and(|value| value.contains('\0'))
        {
            return Err(Error::NulString);
        }
        if let Some(effect) = &self.effect {
            effect.validate()?;
        }
        self.animation_parameters.validate()?;
        self.custom_parameters.validate()
    }
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{AnimationParameters, CustomParameters, Effect, Settings};
    use crate::Error;

    #[test]
    fn effects_map_native_identifiers_losslessly() {
        for effect in [
            Effect::None,
            Effect::Dissolve,
            Effect::MagicMove,
            Effect::Unknown("com.example.future".into()),
        ] {
            assert_eq!(Effect::from_identifier(effect.identifier()), effect);
        }
    }

    #[test]
    fn known_identifiers_cannot_be_smuggled_as_unknown_values() {
        assert!(!Effect::Unknown("none".into()).is_canonical());
        assert!(!Effect::Unknown("apple:dissolve".into()).is_canonical());
        assert!(!Effect::Unknown("apple:magic-move-implied-motion-path".into()).is_canonical());
        assert!(Effect::Unknown("com.example.future".into()).is_canonical());
    }

    #[test]
    fn transition_owned_handles_are_compact() {
        assert_eq!(
            size_of::<Option<Box<[u8]>>>(),
            size_of::<Option<Box<str>>>(),
        );
        assert_eq!(size_of::<Option<Box<[u8]>>>(), 2 * size_of::<usize>(),);
        assert_eq!(size_of::<CustomParameters>() % 4, 0);
    }

    #[test]
    fn transition_unknown_values_and_opaque_payloads_are_lossless() {
        let settings = Settings {
            animation_type: Some("future-transition".into()),
            effect: Some(Effect::from_identifier("com.example.future")),
            animation_parameters: AnimationParameters {
                color_payload: Some(vec![0xff, 0x00, 0x7f].into_boxed_slice()),
                timing_curve_payloads: [None, Some(vec![0xff].into_boxed_slice()), None],
                timing_curve_theme_names: [Some("Future Curve".into()), None, None],
                ..AnimationParameters::default()
            },
            custom_parameters: CustomParameters {
                acceleration: Some(super::Acceleration::from_native(99)),
                text_delivery: Some(super::TextDelivery::from_native(-7)),
                mosaic_type: Some(super::MosaicType::from_native(u32::MAX)),
                ..CustomParameters::default()
            },
            ..Settings::default()
        };

        assert_eq!(settings.validate(), Ok(()));
        assert_eq!(
            settings.effect.as_ref().map(Effect::identifier),
            Some("com.example.future")
        );
        assert_eq!(
            settings.animation_parameters.color_payload(),
            Some(&[0xff, 0x00, 0x7f][..])
        );
        assert_eq!(
            settings.animation_parameters.timing_curve_payloads()[1].as_deref(),
            Some(&[0xff][..])
        );
        assert_eq!(
            settings.animation_parameters.timing_curve_theme_names()[0].as_deref(),
            Some("Future Curve")
        );
    }

    #[test]
    fn transition_validation_reports_typed_failures() {
        let mut settings = Settings {
            duration: Some(-0.1),
            ..Settings::default()
        };
        assert_eq!(settings.validate(), Err(Error::InvalidDuration));

        settings.duration = None;
        settings.delay = Some(f64::NAN);
        assert_eq!(settings.validate(), Err(Error::InvalidDelay));

        settings.delay = None;
        settings.custom_parameters.twist = Some(f32::INFINITY);
        assert_eq!(settings.validate(), Err(Error::InvalidCustomFloat));

        settings.custom_parameters.twist = None;
        settings.animation_parameters.detail = Some(f64::NEG_INFINITY);
        assert_eq!(settings.validate(), Err(Error::InvalidDetail));

        settings.animation_parameters.detail = None;
        settings.animation_type = Some("bad\0name".into());
        assert_eq!(settings.validate(), Err(Error::NulString));

        settings.animation_type = None;
        settings.effect = Some(Effect::Unknown("none".into()));
        assert_eq!(settings.validate(), Err(Error::NonCanonicalEffect));

        settings.effect = None;
        settings.animation_parameters.timing_curve_theme_names[2] = Some("bad\0name".into());
        assert_eq!(settings.validate(), Err(Error::NulString));
    }

    #[test]
    fn effect_constructors_validate_canonical_and_nul_rules() {
        assert_eq!(Effect::unknown("none"), Err(Error::NonCanonicalEffect));
        assert_eq!(Effect::unknown("future\0effect"), Err(Error::NulString));
        assert_eq!(
            Effect::unknown("com.example.future")
                .as_ref()
                .map(Effect::identifier),
            Ok("com.example.future")
        );
    }

    #[test]
    fn transition_scalars_are_compact_and_lossless() {
        assert_eq!(size_of::<super::Direction>(), 4);
        assert_eq!(size_of::<super::MosaicType>(), 4);
        assert_eq!(size_of::<super::Acceleration>(), 4);
        assert_eq!(size_of::<super::TextDelivery>(), 4);
        assert_eq!(super::Acceleration::EaseInOut.native_value(), 4);
        assert_eq!(super::Acceleration::from_native(19).native_value(), 19);
        assert_eq!(
            super::Acceleration::from_native(i32::MIN).native_value(),
            i32::MIN
        );
        assert_eq!(
            super::Acceleration::from_native(i32::MAX).native_value(),
            i32::MAX
        );
        assert_eq!(super::TextDelivery::from_native(-1).native_value(), -1);
        assert_eq!(
            super::TextDelivery::from_native(i32::MAX).native_value(),
            i32::MAX
        );
        assert_eq!(
            super::Direction::from_native(u32::MAX).native_value(),
            u32::MAX
        );
        assert_eq!(super::Acceleration::from_native(19).kind(), None);
        assert_eq!(
            super::TextDelivery::ByCharacter.kind(),
            Some(super::TextDeliveryKind::ByCharacter)
        );
    }
}
