use std::error::Error;
use std::fmt;

/// Hard ceiling for native IWA objects indexed by one Keynote package.
pub const MAX_OBJECTS: usize = 1_000_000;
/// Hard ceiling for semantic slides decoded from one Keynote package.
pub const MAX_SLIDES: usize = 65_536;
/// Hard ceiling for semantic graph-reference occurrences traversed by one package.
pub const MAX_REFERENCES: usize = 1_000_000;
/// Hard ceiling for text-storage objects decoded from one Keynote package.
pub const MAX_TEXT_STORAGES: usize = 1_000_000;
/// Hard ceiling for rich-text fragment ranges retained by one Keynote package.
pub const MAX_TEXT_FRAGMENTS: usize = 1_000_000;
/// Hard ceiling for aggregate decoded text bytes in one Keynote package.
pub const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// A semantic Keynote resource governed by [`SemanticLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Native IWA objects indexed for package-wide lookup.
    Objects,
    /// Slides decoded from the native show tree.
    Slides,
    /// Semantic graph-reference occurrences traversed from the show root.
    References,
    /// Native text-storage objects decoded by the semantic reader.
    TextStorages,
    /// Rich-text fragment ranges retained by the semantic reader.
    TextFragments,
    /// Aggregate bytes retained from text storage and semantic identifiers.
    TextBytes,
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Objects => "objects",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
        })
    }
}

/// An invalid caller-selected semantic resource ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct SemanticLimitsError {
    /// Resource category whose requested limit is invalid.
    pub kind: SemanticLimitKind,
    /// Requested resource ceiling.
    pub value: usize,
    /// Format-wide hard ceiling for the resource.
    pub maximum: usize,
}

impl fmt::Display for SemanticLimitsError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "Keynote semantic {} limit must be non-zero and no greater than {}, got {}",
            self.kind, self.maximum, self.value
        )
    }
}

impl Error for SemanticLimitsError {}

/// Checked resource ceilings for semantic decoding of one Keynote package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    objects: usize,
    slides: usize,
    references: usize,
    text_storages: usize,
    text_fragments: usize,
    text_bytes: usize,
}

impl SemanticLimits {
    /// Hard ceiling for native IWA objects indexed by one Keynote package.
    pub const MAX_OBJECTS: usize = MAX_OBJECTS;
    /// Hard ceiling for semantic slides decoded from one Keynote package.
    pub const MAX_SLIDES: usize = MAX_SLIDES;
    /// Hard ceiling for semantic graph-reference occurrences.
    pub const MAX_REFERENCES: usize = MAX_REFERENCES;
    /// Hard ceiling for decoded text-storage objects in one Keynote package.
    pub const MAX_TEXT_STORAGES: usize = MAX_TEXT_STORAGES;
    /// Hard ceiling for retained rich-text fragment ranges.
    pub const MAX_TEXT_FRAGMENTS: usize = MAX_TEXT_FRAGMENTS;
    /// Hard ceiling for aggregate decoded text bytes in one Keynote package.
    pub const MAX_TEXT_BYTES: usize = MAX_TEXT_BYTES;

    /// Build a checked semantic resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error when any requested ceiling is zero or exceeds its
    /// format-wide hard ceiling.
    pub const fn new(
        max_objects: usize,
        max_slides: usize,
        max_references: usize,
        max_text_storages: usize,
        max_text_fragments: usize,
        max_text_bytes: usize,
    ) -> Result<Self, SemanticLimitsError> {
        if max_objects == 0 || max_objects > MAX_OBJECTS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Objects,
                value: max_objects,
                maximum: MAX_OBJECTS,
            });
        }
        if max_slides == 0 || max_slides > MAX_SLIDES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Slides,
                value: max_slides,
                maximum: MAX_SLIDES,
            });
        }
        if max_references == 0 || max_references > MAX_REFERENCES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::References,
                value: max_references,
                maximum: MAX_REFERENCES,
            });
        }
        if max_text_storages == 0 || max_text_storages > MAX_TEXT_STORAGES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::TextStorages,
                value: max_text_storages,
                maximum: MAX_TEXT_STORAGES,
            });
        }
        if max_text_fragments == 0 || max_text_fragments > MAX_TEXT_FRAGMENTS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::TextFragments,
                value: max_text_fragments,
                maximum: MAX_TEXT_FRAGMENTS,
            });
        }
        if max_text_bytes == 0 || max_text_bytes > MAX_TEXT_BYTES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::TextBytes,
                value: max_text_bytes,
                maximum: MAX_TEXT_BYTES,
            });
        }
        Ok(Self {
            objects: max_objects,
            slides: max_slides,
            references: max_references,
            text_storages: max_text_storages,
            text_fragments: max_text_fragments,
            text_bytes: max_text_bytes,
        })
    }

    /// Maximum number of native IWA objects indexed for package-wide lookup.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.objects
    }

    /// Maximum number of semantic slides decoded from the native show tree.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.slides
    }

    /// Maximum semantic graph-reference occurrences traversed.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.references
    }

    /// Maximum number of native text-storage objects decoded.
    #[must_use]
    pub const fn max_text_storages(self) -> usize {
        self.text_storages
    }

    /// Maximum rich-text fragment ranges retained.
    #[must_use]
    pub const fn max_text_fragments(self) -> usize {
        self.text_fragments
    }

    /// Maximum aggregate byte length of retained semantic text and identifiers.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            objects: MAX_OBJECTS,
            slides: MAX_SLIDES,
            references: MAX_REFERENCES,
            text_storages: MAX_TEXT_STORAGES,
            text_fragments: MAX_TEXT_FRAGMENTS,
            text_bytes: MAX_TEXT_BYTES,
        }
    }
}

/// Physical and semantic resource profiles used to read a Keynote package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReadOptions {
    archive: litchi_iwa_archive::Limits,
    semantic: SemanticLimits,
}

impl ReadOptions {
    /// Combine checked physical and semantic resource profiles.
    #[must_use]
    pub const fn new(archive: litchi_iwa_archive::Limits, semantic: SemanticLimits) -> Self {
        Self { archive, semantic }
    }

    /// Return the physical archive ingress profile.
    #[must_use]
    pub const fn archive(self) -> litchi_iwa_archive::Limits {
        self.archive
    }

    /// Return the semantic decoding profile.
    #[must_use]
    pub const fn semantic(self) -> SemanticLimits {
        self.semantic
    }
}

impl Default for ReadOptions {
    fn default() -> Self {
        Self::new(
            litchi_iwa_archive::Limits::default(),
            SemanticLimits::default(),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::{
        MAX_OBJECTS, MAX_REFERENCES, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS,
        MAX_TEXT_STORAGES, ReadOptions, SemanticLimitKind, SemanticLimits,
    };

    #[test]
    fn defaults_use_every_hard_ceiling() {
        let limits = SemanticLimits::default();
        assert_eq!(limits.max_objects(), MAX_OBJECTS);
        assert_eq!(limits.max_slides(), MAX_SLIDES);
        assert_eq!(limits.max_references(), MAX_REFERENCES);
        assert_eq!(limits.max_text_storages(), MAX_TEXT_STORAGES);
        assert_eq!(limits.max_text_fragments(), MAX_TEXT_FRAGMENTS);
        assert_eq!(limits.max_text_bytes(), MAX_TEXT_BYTES);
    }

    #[test]
    fn checked_profile_retains_each_limit() {
        let result = SemanticLimits::new(1, 2, 3, 4, 5, 6);
        let Ok(limits) = result else {
            panic!("positive limits below the hard ceilings should be valid");
        };
        assert_eq!(limits.max_objects(), 1);
        assert_eq!(limits.max_slides(), 2);
        assert_eq!(limits.max_references(), 3);
        assert_eq!(limits.max_text_storages(), 4);
        assert_eq!(limits.max_text_fragments(), 5);
        assert_eq!(limits.max_text_bytes(), 6);
    }

    #[test]
    fn zero_and_over_cap_limits_are_rejected_with_resource_details() {
        let invalid = [
            (
                SemanticLimits::new(0, 1, 1, 1, 1, 1),
                SemanticLimitKind::Objects,
                0,
                MAX_OBJECTS,
            ),
            (
                SemanticLimits::new(1, MAX_SLIDES + 1, 1, 1, 1, 1),
                SemanticLimitKind::Slides,
                MAX_SLIDES + 1,
                MAX_SLIDES,
            ),
            (
                SemanticLimits::new(1, 1, 0, 1, 1, 1),
                SemanticLimitKind::References,
                0,
                MAX_REFERENCES,
            ),
            (
                SemanticLimits::new(1, 1, 1, 0, 1, 1),
                SemanticLimitKind::TextStorages,
                0,
                MAX_TEXT_STORAGES,
            ),
            (
                SemanticLimits::new(1, 1, 1, 1, 0, 1),
                SemanticLimitKind::TextFragments,
                0,
                MAX_TEXT_FRAGMENTS,
            ),
            (
                SemanticLimits::new(1, 1, 1, 1, 1, MAX_TEXT_BYTES + 1),
                SemanticLimitKind::TextBytes,
                MAX_TEXT_BYTES + 1,
                MAX_TEXT_BYTES,
            ),
        ];

        for (result, expected_kind, expected_value, expected_maximum) in invalid {
            let Err(error) = result else {
                panic!("an invalid semantic ceiling should be rejected");
            };
            assert_eq!(error.kind, expected_kind);
            assert_eq!(error.value, expected_value);
            assert_eq!(error.maximum, expected_maximum);
        }
    }

    #[test]
    fn read_options_retain_both_profiles() {
        let archive = litchi_iwa_archive::Limits::default();
        let semantic_result = SemanticLimits::new(6, 5, 4, 3, 2, 1);
        let Ok(semantic) = semantic_result else {
            panic!("positive limits below the hard ceilings should be valid");
        };
        let options = ReadOptions::new(archive, semantic);
        assert_eq!(options.archive(), archive);
        assert_eq!(options.semantic(), semantic);
    }
}
