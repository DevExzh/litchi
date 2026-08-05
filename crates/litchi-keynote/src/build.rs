//! Keynote build-animation semantic values.

use crate::{Error, Result, Seconds};

mod semantic;

pub use semantic::{
    Action, Blink, Bounce, Effect, Emphasis, Finite, Flip, FlipDirection, HorizontalDirection,
    Jiggle, JiggleIntensity, Keyboard, KeyboardDirection, MAX_IDENTIFIER_BYTES, MAX_PATH_NODES,
    MAX_PATH_SUBPATHS, Motion, Node, NodeKind, ObjectEffect, Opacity, Path, Point, Pop, Pulse,
    Rotation, RotationDirection, Scale, Settings, Subpath, SwooshDirection, TimingCurve,
    UnknownText,
};

const NONE_ACCELERATION: i32 = 0;
const EASE_IN_ACCELERATION: i32 = 1;
const EASE_OUT_ACCELERATION: i32 = 2;
const EASE_IN_OUT_ACCELERATION: i32 = 3;
const CUSTOM_ACCELERATION: i32 = 4;

/// The relationship between one build event and the preceding presentation
/// event.
#[repr(u8)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Start {
    /// Advance to this build with a presenter click.
    OnClick,
    /// Start automatically after the slide transition.
    AfterTransition,
    /// Start concurrently with the preceding build event.
    WithPrevious,
    /// Start after the preceding build event completes.
    AfterPrevious,
}

/// Recognized timing curves for a build action.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum AccelerationKind {
    /// Constant speed.
    None,
    /// Ease into the action.
    EaseIn,
    /// Ease out of the action.
    EaseOut,
    /// Ease into and out of the action.
    EaseInOut,
    /// Use a custom timing curve.
    Custom,
}

/// The compact native timing-curve value used by build actions.
///
/// Known curves have named associated constants. Future native values are
/// retained losslessly so an adapter can read and write them unchanged.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Acceleration(i32);

impl Acceleration {
    /// Constant-speed action.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const None: Self = Self(NONE_ACCELERATION);
    /// Ease into the action.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseIn: Self = Self(EASE_IN_ACCELERATION);
    /// Ease out of the action.
    #[allow(
        non_upper_case_globals,
        reason = "enum-style associated constants are the ergonomic public API"
    )]
    pub const EaseOut: Self = Self(EASE_OUT_ACCELERATION);
    /// Ease into and out of the action.
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

    /// Wrap a native build timing-curve value.
    #[must_use]
    pub const fn from_native(value: i32) -> Self {
        Self(value)
    }

    /// Return the native build timing-curve value.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.0
    }

    /// Return the recognized timing-curve kind, if known.
    #[must_use]
    pub const fn kind(self) -> Option<AccelerationKind> {
        match self.0 {
            NONE_ACCELERATION => Some(AccelerationKind::None),
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

/// A lossless semantic build effect identifier.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum AnimationType {
    /// An object appears.
    Appear,
    /// An object dissolves.
    Dissolve,
    /// An object moves into view.
    MoveIn,
    /// An object scales into view.
    Scale,
    /// An object fades and scales into view.
    FadeAndScale,
    /// A producer-specific effect not yet modeled by this crate.
    Unknown(String),
}

impl AnimationType {
    /// Decode a producer identifier while retaining unknown values.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyIdentifier`] when `identifier` is empty.
    pub fn from_identifier(identifier: &str) -> Result<Self> {
        if identifier.is_empty() {
            return Err(Error::EmptyIdentifier);
        }

        let effect = if identifier.eq_ignore_ascii_case("appear") {
            Self::Appear
        } else if identifier.eq_ignore_ascii_case("dissolve") {
            Self::Dissolve
        } else if contains_ascii_case_insensitive(identifier, b"move") {
            Self::MoveIn
        } else if contains_ascii_case_insensitive(identifier, b"fade")
            && contains_ascii_case_insensitive(identifier, b"scale")
        {
            Self::FadeAndScale
        } else if contains_ascii_case_insensitive(identifier, b"scale") {
            Self::Scale
        } else {
            Self::Unknown(identifier.to_owned())
        };
        Ok(effect)
    }

    /// Return the canonical semantic label or the preserved producer value.
    #[must_use]
    pub fn identifier(&self) -> &str {
        match self {
            Self::Appear => "appear",
            Self::Dissolve => "dissolve",
            Self::MoveIn => "move",
            Self::Scale => "scale",
            Self::FadeAndScale => "fade-scale",
            Self::Unknown(identifier) => identifier,
        }
    }

    /// Return a stable display name.
    #[must_use]
    pub fn name(&self) -> &str {
        match self {
            Self::Appear => "Appear",
            Self::Dissolve => "Dissolve",
            Self::MoveIn => "Move In",
            Self::Scale => "Scale",
            Self::FadeAndScale => "Fade and Scale",
            Self::Unknown(identifier) => identifier,
        }
    }
}

/// A build animation attached to a semantic slide.
#[derive(Debug, Clone, PartialEq)]
pub struct Build {
    animation_type: AnimationType,
    duration: Seconds,
}

impl Build {
    /// Construct a build from its validated semantic values.
    #[must_use]
    pub const fn new(animation_type: AnimationType, duration: Seconds) -> Self {
        Self {
            animation_type,
            duration,
        }
    }

    /// Return the effect kind.
    #[must_use]
    pub const fn animation_type(&self) -> &AnimationType {
        &self.animation_type
    }

    /// Return the build duration.
    #[must_use]
    pub const fn duration(&self) -> Seconds {
        self.duration
    }
}

fn contains_ascii_case_insensitive(haystack: &str, needle: &[u8]) -> bool {
    haystack
        .as_bytes()
        .windows(needle.len())
        .any(|window| window.eq_ignore_ascii_case(needle))
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::mem::size_of;

    #[test]
    fn scalar_values_are_compact_and_lossless() {
        assert_eq!(size_of::<Start>(), 1);
        assert_eq!(size_of::<Acceleration>(), size_of::<i32>());
        assert_eq!(Acceleration::None.kind(), Some(AccelerationKind::None));
        assert_eq!(Acceleration::EaseInOut.native_value(), 3);

        let unknown = Acceleration::from_native(99);
        assert_eq!(unknown.native_value(), 99);
        assert_eq!(unknown.kind(), None);
    }

    #[test]
    fn identifiers_are_lossless_and_typed() -> Result<()> {
        assert_eq!(
            AnimationType::from_identifier("appear")?,
            AnimationType::Appear
        );
        let unknown = AnimationType::from_identifier("com.example.future")?;
        assert_eq!(unknown.identifier(), "com.example.future");
        assert_eq!(unknown.name(), "com.example.future");
        Ok(())
    }

    #[test]
    fn identifier_matching_is_case_insensitive_without_changing_unknown_bytes() -> Result<()> {
        assert_eq!(
            AnimationType::from_identifier("APPEAR")?,
            AnimationType::Appear
        );
        assert_eq!(
            AnimationType::from_identifier("Dissolve")?,
            AnimationType::Dissolve
        );
        assert_eq!(
            AnimationType::from_identifier("custom-MoVe")?,
            AnimationType::MoveIn
        );
        assert_eq!(
            AnimationType::from_identifier("FADE-and-SCALE")?,
            AnimationType::FadeAndScale
        );
        assert_eq!(
            AnimationType::from_identifier("future-SCALE")?,
            AnimationType::Scale
        );

        let unknown = "Future-Éffect";
        let parsed = AnimationType::from_identifier(unknown)?;
        assert_eq!(parsed, AnimationType::Unknown(unknown.to_owned()));
        assert_eq!(parsed.identifier(), unknown);
        Ok(())
    }

    #[test]
    fn invalid_animation_inputs_are_rejected() {
        assert_eq!(Seconds::new(f64::NAN), Err(Error::InvalidDuration));
        assert_eq!(
            AnimationType::from_identifier(""),
            Err(Error::EmptyIdentifier)
        );
    }
}
