use std::error::Error;
use std::fmt;

use crate::{MAX_SHEETS as HARD_MAX_SHEETS, MAX_TABLES as HARD_MAX_TABLES};

/// Hard ceiling for native IWA objects indexed by one Numbers package.
pub const MAX_OBJECTS: usize = 1_000_000;
/// Hard ceiling for the caller-selected reference setting.
///
/// The selected value independently limits rooted graph-reference occurrences
/// during strict construction and unique source-derived entries retained by
/// one formula-enrichment build. Formula wire, work, text, and depth limits are
/// fixed adapter-owned safeguards.
pub const MAX_REFERENCES: usize = 1_000_000;

/// A Numbers resource reported by a semantic-limit error.
///
/// `Objects`, `Sheets`, `Tables`, and `References` can be selected through
/// [`SemanticLimits`]. Formula-scan resources are fixed adapter-owned ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Native IWA objects retained in the package-local lookup index.
    Objects,
    /// Sheets traversed from the Numbers document root.
    Sheets,
    /// Tables retained by either semantic projection.
    Tables,
    /// Caller-selected rooted graph references or unique source-derived
    /// retained formula-enrichment entries.
    References,
    /// Fixed aggregate bytes of type-6383 category metadata scanned by formula
    /// enrichment.
    FormulaWireBytes,
    /// Fixed aggregate formula-discovery and category-preflight work.
    FormulaWork,
    /// Fixed aggregate source text admitted to formula enrichment.
    TextBytes,
    /// Fixed nesting depth of one formula-category tree.
    FormulaDepth,
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Objects => "objects",
            Self::Sheets => "sheets",
            Self::Tables => "tables",
            Self::References => "references",
            Self::FormulaWireBytes => "formula wire bytes",
            Self::FormulaWork => "formula work",
            Self::TextBytes => "text bytes",
            Self::FormulaDepth => "formula depth",
        })
    }
}

/// An invalid caller-selected semantic resource ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct SemanticLimitsError {
    /// Resource category whose requested ceiling is invalid.
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
            "Numbers semantic {} limit must be non-zero and no greater than {}, got {}",
            self.kind, self.maximum, self.value
        )
    }
}

impl Error for SemanticLimitsError {}

/// Checked resource ceilings for semantic decoding of one Numbers package.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    objects: usize,
    sheets: usize,
    tables: usize,
    references: usize,
}

impl SemanticLimits {
    /// Hard ceiling for indexed native objects.
    pub const MAX_OBJECTS: usize = MAX_OBJECTS;
    /// Hard ceiling for rooted semantic sheets.
    pub const MAX_SHEETS: usize = HARD_MAX_SHEETS;
    /// Hard ceiling for semantic tables.
    pub const MAX_TABLES: usize = HARD_MAX_TABLES;
    /// Hard ceiling for the caller-selected reference setting.
    ///
    /// The selected value independently limits rooted graph-reference
    /// occurrences during strict construction and unique source-derived
    /// entries retained by one formula-enrichment build.
    pub const MAX_REFERENCES: usize = MAX_REFERENCES;

    /// Build a checked semantic resource profile.
    ///
    /// # Errors
    ///
    /// Returns an error when any requested ceiling is zero or exceeds its
    /// format-wide hard ceiling.
    pub const fn new(
        max_objects: usize,
        max_sheets: usize,
        max_tables: usize,
        max_references: usize,
    ) -> Result<Self, SemanticLimitsError> {
        if max_objects == 0 || max_objects > MAX_OBJECTS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Objects,
                value: max_objects,
                maximum: MAX_OBJECTS,
            });
        }
        if max_sheets == 0 || max_sheets > HARD_MAX_SHEETS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Sheets,
                value: max_sheets,
                maximum: HARD_MAX_SHEETS,
            });
        }
        if max_tables == 0 || max_tables > HARD_MAX_TABLES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::Tables,
                value: max_tables,
                maximum: HARD_MAX_TABLES,
            });
        }
        if max_references == 0 || max_references > MAX_REFERENCES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::References,
                value: max_references,
                maximum: MAX_REFERENCES,
            });
        }
        Ok(Self {
            objects: max_objects,
            sheets: max_sheets,
            tables: max_tables,
            references: max_references,
        })
    }

    /// Maximum number of indexed native objects.
    #[must_use]
    pub const fn max_objects(self) -> usize {
        self.objects
    }

    /// Maximum number of rooted semantic sheets.
    #[must_use]
    pub const fn max_sheets(self) -> usize {
        self.sheets
    }

    /// Maximum number of semantic tables retained by either projection.
    #[must_use]
    pub const fn max_tables(self) -> usize {
        self.tables
    }

    /// Caller-selected cap independently applied to rooted references and
    /// unique source-derived retained formula-enrichment entries.
    #[must_use]
    pub const fn max_references(self) -> usize {
        self.references
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            objects: MAX_OBJECTS,
            sheets: HARD_MAX_SHEETS,
            tables: HARD_MAX_TABLES,
            references: MAX_REFERENCES,
        }
    }
}

/// Physical and semantic resource profiles used to read a Numbers package.
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

    /// Return the physical archive-ingress profile.
    #[must_use]
    pub const fn archive(self) -> litchi_iwa_archive::Limits {
        self.archive
    }

    /// Return the semantic-decoding profile.
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
    use super::*;

    #[test]
    fn semantic_limits_reject_zero_and_relaxed_hard_ceilings() {
        assert!(matches!(
            SemanticLimits::new(0, HARD_MAX_SHEETS, HARD_MAX_TABLES, MAX_REFERENCES),
            Err(SemanticLimitsError {
                kind: SemanticLimitKind::Objects,
                ..
            })
        ));
        assert!(matches!(
            SemanticLimits::new(
                MAX_OBJECTS,
                HARD_MAX_SHEETS,
                HARD_MAX_TABLES + 1,
                MAX_REFERENCES,
            ),
            Err(SemanticLimitsError {
                kind: SemanticLimitKind::Tables,
                ..
            })
        ));
    }

    #[test]
    fn default_read_options_preserve_both_profiles() {
        let options = ReadOptions::default();
        assert_eq!(options.archive(), litchi_iwa_archive::Limits::default());
        assert_eq!(options.semantic(), SemanticLimits::default());
    }
}
