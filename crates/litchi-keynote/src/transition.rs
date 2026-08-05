//! Keynote transition vocabulary independent of archive and wire models.

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
    Unknown(String),
}

impl Effect {
    /// Decode a native effect identifier without discarding unknown values.
    #[must_use]
    pub fn from_identifier(identifier: &str) -> Self {
        match identifier {
            NONE_EFFECT => Self::None,
            DISSOLVE_EFFECT => Self::Dissolve,
            MAGIC_MOVE_EFFECT => Self::MagicMove,
            other => Self::Unknown(other.to_owned()),
        }
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
                    identifier.as_str(),
                    NONE_EFFECT | DISSOLVE_EFFECT | MAGIC_MOVE_EFFECT
                )
        )
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

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::Effect;

    #[test]
    fn effects_map_native_identifiers_losslessly() {
        for effect in [
            Effect::None,
            Effect::Dissolve,
            Effect::MagicMove,
            Effect::Unknown("com.example.future".to_owned()),
        ] {
            assert_eq!(Effect::from_identifier(effect.identifier()), effect);
        }
    }

    #[test]
    fn known_identifiers_cannot_be_smuggled_as_unknown_values() {
        assert!(!Effect::Unknown("none".to_owned()).is_canonical());
        assert!(!Effect::Unknown("apple:dissolve".to_owned()).is_canonical());
        assert!(!Effect::Unknown("apple:magic-move-implied-motion-path".to_owned()).is_canonical());
        assert!(Effect::Unknown("com.example.future".to_owned()).is_canonical());
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
