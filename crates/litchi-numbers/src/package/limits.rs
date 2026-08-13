use std::error::Error;
use std::fmt;

use crate::{
    DEFAULT_MAX_TEXT_BYTES, MAX_MATERIALIZED_CELLS, MAX_SHEETS as HARD_MAX_SHEETS,
    MAX_TABLES as HARD_MAX_TABLES,
};

/// Hard ceiling for native IWA objects indexed by one Numbers package.
pub const MAX_OBJECTS: usize = 1_000_000;
/// Hard ceiling for the caller-selected reference setting.
///
/// The selected value independently limits rooted graph-reference occurrences
/// during strict construction and unique source-derived entries retained by
/// one formula-enrichment build. Formula wire, work, text, and depth limits are
/// fixed adapter-owned safeguards.
pub const MAX_REFERENCES: usize = 1_000_000;
/// Hard ceiling for materialized cells retained across one package projection.
const MAX_PROJECTED_CELLS: usize = MAX_MATERIALIZED_CELLS;
/// Hard ceiling for UTF-8 text retained across one package projection.
const MAX_OUTPUT_TEXT_BYTES: usize = DEFAULT_MAX_TEXT_BYTES;
/// Hard ceiling for formula AST nodes processed across one package projection.
const MAX_FORMULA_RENDER_WORK: usize = MAX_REFERENCES;
/// Hard ceiling for nested formula thunk arrays.
const MAX_FORMULA_RENDER_DEPTH: usize = 64;

/// A Numbers resource reported by a semantic-limit error.
///
/// Every resource can be selected through [`SemanticLimits`]. Formula metadata
/// scanning retains additional fixed adapter-owned ceilings.
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
    /// Materialized cells retained across all projected tables.
    MaterializedCells,
    /// UTF-8 bytes retained in table names and textual cell values.
    OutputTextBytes,
    /// Formula AST nodes processed across all projected tables.
    FormulaRenderWork,
    /// Nested formula thunk arrays processed by one formula.
    FormulaRenderDepth,
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
            Self::MaterializedCells => "materialized cells",
            Self::OutputTextBytes => "output text bytes",
            Self::FormulaRenderWork => "formula render work",
            Self::FormulaRenderDepth => "formula render depth",
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
    materialized_cells: usize,
    output_text_bytes: usize,
    formula_render_work: usize,
    formula_render_depth: usize,
}

impl SemanticLimits {
    pub(crate) const fn for_document(limits: crate::DocumentLimits) -> Self {
        Self {
            objects: MAX_OBJECTS,
            sheets: if limits.max_sheets() < HARD_MAX_SHEETS {
                limits.max_sheets()
            } else {
                HARD_MAX_SHEETS
            },
            tables: if limits.max_tables() < HARD_MAX_TABLES {
                limits.max_tables()
            } else {
                HARD_MAX_TABLES
            },
            references: MAX_REFERENCES,
            materialized_cells: if limits.max_materialized_cells() < MAX_PROJECTED_CELLS {
                limits.max_materialized_cells()
            } else {
                MAX_PROJECTED_CELLS
            },
            output_text_bytes: if limits.max_text_bytes() < MAX_OUTPUT_TEXT_BYTES {
                limits.max_text_bytes()
            } else {
                MAX_OUTPUT_TEXT_BYTES
            },
            formula_render_work: MAX_FORMULA_RENDER_WORK,
            formula_render_depth: MAX_FORMULA_RENDER_DEPTH,
        }
    }

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
    /// Hard ceiling for materialized cells across all projected tables.
    pub const MAX_MATERIALIZED_CELLS: usize = MAX_PROJECTED_CELLS;
    /// Hard ceiling for retained semantic UTF-8 text.
    pub const MAX_OUTPUT_TEXT_BYTES: usize = MAX_OUTPUT_TEXT_BYTES;
    /// Hard ceiling for formula AST render work.
    pub const MAX_FORMULA_RENDER_WORK: usize = MAX_FORMULA_RENDER_WORK;
    /// Hard ceiling for nested formula thunk arrays.
    pub const MAX_FORMULA_RENDER_DEPTH: usize = MAX_FORMULA_RENDER_DEPTH;

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
            materialized_cells: MAX_PROJECTED_CELLS,
            output_text_bytes: MAX_OUTPUT_TEXT_BYTES,
            formula_render_work: MAX_FORMULA_RENDER_WORK,
            formula_render_depth: MAX_FORMULA_RENDER_DEPTH,
        })
    }

    /// Select package-wide materialized-cell and retained-text ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error when either ceiling is zero or exceeds its format-wide
    /// hard maximum.
    pub const fn with_projection_limits(
        mut self,
        max_materialized_cells: usize,
        max_output_text_bytes: usize,
    ) -> Result<Self, SemanticLimitsError> {
        if max_materialized_cells == 0 || max_materialized_cells > MAX_PROJECTED_CELLS {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::MaterializedCells,
                value: max_materialized_cells,
                maximum: MAX_PROJECTED_CELLS,
            });
        }
        if max_output_text_bytes == 0 || max_output_text_bytes > MAX_OUTPUT_TEXT_BYTES {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::OutputTextBytes,
                value: max_output_text_bytes,
                maximum: MAX_OUTPUT_TEXT_BYTES,
            });
        }
        self.materialized_cells = max_materialized_cells;
        self.output_text_bytes = max_output_text_bytes;
        Ok(self)
    }

    /// Select package-wide formula-render work and nesting ceilings.
    ///
    /// # Errors
    ///
    /// Returns an error when either ceiling is zero or exceeds its format-wide
    /// hard maximum.
    pub const fn with_formula_render_limits(
        mut self,
        max_work: usize,
        max_depth: usize,
    ) -> Result<Self, SemanticLimitsError> {
        if max_work == 0 || max_work > MAX_FORMULA_RENDER_WORK {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::FormulaRenderWork,
                value: max_work,
                maximum: MAX_FORMULA_RENDER_WORK,
            });
        }
        if max_depth == 0 || max_depth > MAX_FORMULA_RENDER_DEPTH {
            return Err(SemanticLimitsError {
                kind: SemanticLimitKind::FormulaRenderDepth,
                value: max_depth,
                maximum: MAX_FORMULA_RENDER_DEPTH,
            });
        }
        self.formula_render_work = max_work;
        self.formula_render_depth = max_depth;
        Ok(self)
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

    /// Maximum materialized cells retained across all projected tables.
    #[must_use]
    pub const fn max_materialized_cells(self) -> usize {
        self.materialized_cells
    }

    /// Maximum UTF-8 bytes retained across all projected tables.
    #[must_use]
    pub const fn max_output_text_bytes(self) -> usize {
        self.output_text_bytes
    }

    /// Maximum formula AST nodes processed across all projected tables.
    #[must_use]
    pub const fn max_formula_render_work(self) -> usize {
        self.formula_render_work
    }

    /// Maximum nested thunk depth accepted by one formula.
    #[must_use]
    pub const fn max_formula_render_depth(self) -> usize {
        self.formula_render_depth
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self {
            objects: MAX_OBJECTS,
            sheets: HARD_MAX_SHEETS,
            tables: HARD_MAX_TABLES,
            references: MAX_REFERENCES,
            materialized_cells: MAX_PROJECTED_CELLS,
            output_text_bytes: MAX_OUTPUT_TEXT_BYTES,
            formula_render_work: MAX_FORMULA_RENDER_WORK,
            formula_render_depth: MAX_FORMULA_RENDER_DEPTH,
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

    #[test]
    fn projection_and_formula_limits_are_checked_and_preserved() {
        let limits = SemanticLimits::default()
            .with_projection_limits(17, 23)
            .unwrap_or_else(|error| panic!("projection limits were rejected: {error}"))
            .with_formula_render_limits(31, 7)
            .unwrap_or_else(|error| panic!("formula limits were rejected: {error}"));
        assert_eq!(limits.max_materialized_cells(), 17);
        assert_eq!(limits.max_output_text_bytes(), 23);
        assert_eq!(limits.max_formula_render_work(), 31);
        assert_eq!(limits.max_formula_render_depth(), 7);

        assert!(matches!(
            SemanticLimits::default().with_projection_limits(0, 1),
            Err(SemanticLimitsError {
                kind: SemanticLimitKind::MaterializedCells,
                ..
            })
        ));
        assert!(matches!(
            SemanticLimits::default().with_formula_render_limits(1, MAX_FORMULA_RENDER_DEPTH + 1),
            Err(SemanticLimitsError {
                kind: SemanticLimitKind::FormulaRenderDepth,
                ..
            })
        ));
    }
}
