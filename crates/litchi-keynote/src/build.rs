//! Keynote build-animation semantic values.

use crate::{Error, Result, Seconds};

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

        let normalized = identifier.to_ascii_lowercase();
        let effect = if normalized == "appear" {
            Self::Appear
        } else if normalized == "dissolve" {
            Self::Dissolve
        } else if normalized.contains("move") {
            Self::MoveIn
        } else if normalized.contains("fade") && normalized.contains("scale") {
            Self::FadeAndScale
        } else if normalized.contains("scale") {
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

#[cfg(test)]
mod tests {
    use super::*;

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
    fn invalid_animation_inputs_are_rejected() {
        assert_eq!(Seconds::new(f64::NAN), Err(Error::InvalidDuration));
        assert_eq!(
            AnimationType::from_identifier(""),
            Err(Error::EmptyIdentifier)
        );
    }
}
