//! Typed native identifiers for Keynote slide-transition effects.

const NONE_EFFECT: &str = "none";
const DISSOLVE_EFFECT: &str = "apple:dissolve";
const MAGIC_MOVE_EFFECT: &str = "apple:magic-move-implied-motion-path";

/// A slide-transition effect understood by Keynote.
///
/// Named variants cover identifiers verified in native Keynote documents.
/// `Unknown` preserves effects added by future releases without reducing the
/// API back to an untyped string.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub enum KeynoteTransitionEffect {
    /// Keynote's native “No Transition Effect” representation.
    None,
    Dissolve,
    MagicMove,
    /// An identifier introduced by a newer Keynote release.
    Unknown(String),
}

impl KeynoteTransitionEffect {
    /// Decode a native effect identifier without discarding unknown values.
    pub fn from_identifier(identifier: &str) -> Self {
        match identifier {
            NONE_EFFECT => Self::None,
            DISSOLVE_EFFECT => Self::Dissolve,
            MAGIC_MOVE_EFFECT => Self::MagicMove,
            identifier => Self::Unknown(identifier.to_owned()),
        }
    }

    /// Return the native Keynote effect identifier.
    pub fn as_identifier(&self) -> &str {
        match self {
            Self::None => NONE_EFFECT,
            Self::Dissolve => DISSOLVE_EFFECT,
            Self::MagicMove => MAGIC_MOVE_EFFECT,
            Self::Unknown(identifier) => identifier,
        }
    }

    pub(super) fn is_canonical(&self) -> bool {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn transition_effects_map_native_identifiers_losslessly() {
        for effect in [
            KeynoteTransitionEffect::None,
            KeynoteTransitionEffect::Dissolve,
            KeynoteTransitionEffect::MagicMove,
            KeynoteTransitionEffect::Unknown("com.example.future".to_owned()),
        ] {
            assert_eq!(
                KeynoteTransitionEffect::from_identifier(effect.as_identifier()),
                effect
            );
        }
    }
}
