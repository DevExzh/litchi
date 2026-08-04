//! Keynote transition vocabulary independent of archive and wire models.

const NONE_EFFECT: &str = "none";
const DISSOLVE_EFFECT: &str = "apple:dissolve";
const MAGIC_MOVE_EFFECT: &str = "apple:magic-move-implied-motion-path";

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

#[cfg(test)]
mod tests {
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
}
