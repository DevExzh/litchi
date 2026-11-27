//! The resolved ingress profile a package retains: the physical archive
//! limits its catalog was parsed under and the semantic projection limits.

use crate::limits::{OpenOptions, SemanticLimits, SourceLimits};

/// The physical and semantic profiles one package was read under.
///
/// The public envelope is [`OpenOptions`]; this crate-private form keeps
/// the checked physical archive limits the catalog uses, so focused
/// transactions reopen candidates under the exact source profile.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Profile {
    archive: litchi_iwa_archive::Limits,
    semantic: SemanticLimits,
}

impl Profile {
    /// Combine checked physical and semantic resource profiles.
    pub(crate) const fn new(archive: litchi_iwa_archive::Limits, semantic: SemanticLimits) -> Self {
        Self { archive, semantic }
    }

    /// Resolve a public ingress envelope.
    pub(crate) fn from_options(options: OpenOptions) -> super::Result<Self> {
        Ok(Self {
            archive: options.source_limits_value().archive()?,
            semantic: options.semantic_limits_value(),
        })
    }

    /// The public ingress envelope of this profile.
    pub(crate) fn options(self) -> OpenOptions {
        OpenOptions::new()
            .source_limits(SourceLimits::from_archive(self.archive))
            .semantic_limits(self.semantic)
    }

    /// Return the physical archive-ingress profile.
    pub(crate) const fn archive(self) -> litchi_iwa_archive::Limits {
        self.archive
    }

    /// Return the semantic-decoding profile.
    pub(crate) const fn semantic(self) -> SemanticLimits {
        self.semantic
    }
}

impl Default for Profile {
    fn default() -> Self {
        Self::from_options(OpenOptions::new()).unwrap_or(Self {
            archive: litchi_iwa_archive::Limits::default(),
            semantic: SemanticLimits::DEFAULT,
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn default_profile_round_trips_the_default_options() {
        let profile = Profile::default();
        assert_eq!(profile.semantic(), SemanticLimits::default());
        assert_eq!(
            profile.archive().max_input_bytes(),
            SourceLimits::DEFAULT.max_input_bytes()
        );
        assert_eq!(profile.options(), OpenOptions::default());
    }
}
