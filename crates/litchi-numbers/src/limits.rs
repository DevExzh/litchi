//! Checked ingress profiles for Numbers packages and workbooks (ADR 0034 §5).
//!
//! [`SourceLimits`] bounds physical ingress (the captured input, its ZIP
//! members and decoded IWA components); [`SemanticLimits`] bounds the
//! semantic projection (objects, sheets, tables, references, cells, text and
//! formula rendering). [`OpenOptions`] carries both into
//! [`Package::open_with_options`](crate::Package::open_with_options) and the
//! read-only [`Workbook`](crate::Workbook) ingress.
//!
//! Both profiles are built with consuming builders whose `max_<dimension>`
//! setters accept `1..=MAXIMUM`; a value outside that range is this
//! module's [`Error::InvalidLimit`] (ADR 0034 folds it into the owning
//! module's error type). An exceeded ceiling at ingress is a typed limit
//! error naming the dimension.
//!
//! ```
//! use litchi_numbers::{OpenOptions, SemanticLimits, SourceLimits};
//!
//! # fn main() -> Result<(), litchi_numbers::limits::Error> {
//! let source = SourceLimits::builder().max_input_bytes(64 * 1024 * 1024)?.build()?;
//! let semantic = SemanticLimits::builder().max_sheets(16)?.build()?;
//! let options = OpenOptions::new()
//!     .source_limits(source)
//!     .semantic_limits(semantic);
//! assert_eq!(options.source_limits_value().max_input_bytes(), 64 * 1024 * 1024);
//! assert_eq!(options.semantic_limits_value().max_sheets(), 16);
//! # Ok(())
//! # }
//! ```

use std::fmt;

use crate::package::Error as PackageError;

/// A limit outside its valid range, from a [`SourceLimitsBuilder`] or
/// [`SemanticLimitsBuilder`] setter or `build()`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// A caller-selected limit is outside `1..=maximum`.
    #[error("Numbers {kind} limit must be between 1 and {maximum}, got {value}")]
    InvalidLimit {
        /// The dimension being configured.
        kind: LimitKind,
        /// The rejected value.
        value: u64,
        /// The hard ceiling for this dimension.
        maximum: u64,
    },
}

/// Hard ceiling for native IWA objects indexed by one Numbers package.
pub const MAX_OBJECTS: usize = 1_000_000;
/// Hard ceiling for the caller-selected reference setting.
///
/// The selected value independently limits rooted graph-reference occurrences
/// during strict construction and unique source-derived entries retained by
/// one formula-enrichment build. Formula wire, work, text, and depth limits are
/// fixed adapter-owned safeguards.
pub const MAX_REFERENCES: usize = 1_000_000;
/// Hard ceiling for rooted sheets in one workbook.
pub const MAX_SHEETS: usize = 4096;
/// Hard ceiling for tables in one workbook.
pub const MAX_TABLES: usize = 65_536;
/// Hard ceiling for materialized cells retained by one workbook.
pub const MAX_MATERIALIZED_CELLS: usize = 16_000_000;
/// Hard ceiling for UTF-8 bytes retained by names, headers and textual cells.
pub const MAX_TEXT_BYTES: usize = 64 * 1024 * 1024;
/// Hard ceiling for formula AST nodes processed across one projection.
const MAX_FORMULA_RENDER_WORK: usize = MAX_REFERENCES;
/// Hard ceiling for nested formula thunk arrays.
const MAX_FORMULA_RENDER_DEPTH: usize = 64;

/// A physical ingress dimension of [`SourceLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Bytes of the complete captured source.
    InputBytes,
    /// ZIP members (directory records included).
    Members,
    /// Declared uncompressed bytes of one ZIP member.
    EntryBytes,
    /// Aggregate declared uncompressed bytes of all members.
    AggregateBytes,
    /// Decompressed bytes of one IWA component.
    ComponentBytes,
    /// Protobuf fields walked in IWA object headers and focused payload
    /// scans.
    HeaderFields,
}

impl fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Members => "members",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
            Self::HeaderFields => "header fields",
        })
    }
}

/// Checked physical ingress ceilings for one Numbers source. Every
/// dimension defaults to its hard ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: u64,
    members: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
    header_fields: usize,
}

impl SourceLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        input_bytes: litchi_iwa_archive::Limits::MAX_INPUT_BYTES,
        members: litchi_iwa_detect::Limits::HARD_MAX_FILES,
        entry_bytes: litchi_iwa_detect::Limits::HARD_MAX_ENTRY_SIZE,
        aggregate_bytes: litchi_iwa_detect::Limits::HARD_MAX_TOTAL_SIZE,
        component_bytes: litchi_iwa_detect::Limits::HARD_MAX_IWA_STREAM_SIZE,
        header_fields: litchi_iwa_core::Limits::MAX_HEADER_FIELDS,
    };

    /// The default ingress profile: every dimension at its hard ceiling.
    ///
    /// Unlike the 512 MiB input default of the OOXML and ODF profiles, the
    /// Numbers default keeps the 2 GiB package ceiling: the focused
    /// transactions size their work budgets from the physical profile and
    /// precharge up to one component ceiling (512 MiB), which a 512 MiB
    /// input ceiling would refuse.
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// A builder starting from this profile.
    pub const fn to_builder(self) -> SourceLimitsBuilder {
        SourceLimitsBuilder { limits: self }
    }

    /// Maximum bytes of the complete captured source.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum ZIP members, directory records included.
    #[must_use]
    pub const fn max_members(self) -> usize {
        self.members
    }

    /// Maximum declared uncompressed bytes of one ZIP member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    /// Maximum aggregate declared uncompressed bytes.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    /// Maximum decompressed bytes of one IWA component.
    #[must_use]
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    /// Maximum protobuf fields walked in IWA object headers and focused
    /// payload scans.
    #[must_use]
    pub const fn max_header_fields(self) -> usize {
        self.header_fields
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SourceLimitKind) -> u64 {
        match kind {
            SourceLimitKind::InputBytes => self.input_bytes,
            SourceLimitKind::Members => self.members as u64,
            SourceLimitKind::EntryBytes => self.entry_bytes,
            SourceLimitKind::AggregateBytes => self.aggregate_bytes,
            SourceLimitKind::ComponentBytes => self.component_bytes as u64,
            SourceLimitKind::HeaderFields => self.header_fields as u64,
        }
    }

    /// The physical archive profile handed to the package catalog.
    pub(crate) fn archive(self) -> Result<litchi_iwa_archive::Limits, PackageError> {
        let limits = litchi_iwa_archive::Limits::new(
            self.input_bytes,
            self.members,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
        .map_err(PackageError::Archive)?;
        if self.header_fields == Self::MAXIMUM.header_fields {
            return Ok(limits);
        }
        let profile = limits
            .archive_limits()
            .with_header_fields(self.header_fields)
            .map_err(|error| PackageError::Archive(litchi_iwa_archive::Error::Iwa(error)))?;
        limits
            .with_archive_limits(profile)
            .map_err(PackageError::Archive)
    }

    /// The detector profile handed to semantic-only ingress.
    pub(crate) fn detector(self) -> litchi_iwa_detect::Result<litchi_iwa_detect::Limits> {
        litchi_iwa_detect::Limits::new(
            self.input_bytes
                .min(litchi_iwa_detect::Limits::HARD_MAX_INPUT_BYTES),
            self.members,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )
    }

    /// Derive the profile of an archive catalog a coordinator already
    /// parsed (unstable migration-host seam).
    #[cfg(feature = "internal-iwork-source")]
    #[doc(hidden)]
    #[must_use]
    pub fn __from_archive(limits: litchi_iwa_archive::Limits) -> Self {
        Self::from_archive(limits)
    }

    /// Derive the profile of an already-validated archive catalog. Values are
    /// clamped to the hard ceilings, never rejected.
    pub(crate) fn from_archive(limits: litchi_iwa_archive::Limits) -> Self {
        let maximum = Self::MAXIMUM;
        Self {
            input_bytes: limits.max_input_bytes().clamp(1, maximum.input_bytes),
            members: limits.max_entries().clamp(1, maximum.members),
            entry_bytes: limits.max_entry_bytes().clamp(1, maximum.entry_bytes),
            aggregate_bytes: limits.max_total_bytes().clamp(1, maximum.aggregate_bytes),
            component_bytes: limits
                .max_iwa_stream_bytes()
                .clamp(1, maximum.component_bytes),
            header_fields: limits
                .archive_limits()
                .max_header_fields()
                .clamp(1, maximum.header_fields),
        }
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SourceLimits`].
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

fn checked_source(kind: SourceLimitKind, value: u64) -> Result<u64, Error> {
    let maximum = SourceLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind: LimitKind::Source(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

fn checked_source_count(kind: SourceLimitKind, value: usize) -> Result<usize, Error> {
    let wide = u64::try_from(value).unwrap_or(u64::MAX);
    checked_source(kind, wide)?;
    Ok(value)
}

impl SourceLimitsBuilder {
    /// Set the maximum bytes of the complete captured source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_input_bytes(mut self, value: u64) -> Result<Self, Error> {
        self.limits.input_bytes = checked_source(SourceLimitKind::InputBytes, value)?;
        Ok(self)
    }

    /// Set the maximum ZIP members.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_members(mut self, value: usize) -> Result<Self, Error> {
        self.limits.members = checked_source_count(SourceLimitKind::Members, value)?;
        Ok(self)
    }

    /// Set the maximum declared uncompressed bytes of one member.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_entry_bytes(mut self, value: u64) -> Result<Self, Error> {
        self.limits.entry_bytes = checked_source(SourceLimitKind::EntryBytes, value)?;
        Ok(self)
    }

    /// Set the maximum aggregate declared uncompressed bytes.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_aggregate_bytes(mut self, value: u64) -> Result<Self, Error> {
        self.limits.aggregate_bytes = checked_source(SourceLimitKind::AggregateBytes, value)?;
        Ok(self)
    }

    /// Set the maximum decompressed bytes of one IWA component.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_component_bytes(mut self, value: usize) -> Result<Self, Error> {
        self.limits.component_bytes = checked_source_count(SourceLimitKind::ComponentBytes, value)?;
        Ok(self)
    }

    /// Set the maximum protobuf fields walked in IWA object headers and
    /// focused payload scans.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_header_fields(mut self, value: usize) -> Result<Self, Error> {
        self.limits.header_fields = checked_source_count(SourceLimitKind::HeaderFields, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] naming the component dimension when
    /// the physical archive cannot represent the selected component ceiling.
    pub fn build(self) -> Result<SourceLimits, Error> {
        self.limits
            .archive()
            .map_err(|_invalid| Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::ComponentBytes),
                value: self.limits.value(SourceLimitKind::ComponentBytes),
                maximum: SourceLimits::MAXIMUM.value(SourceLimitKind::ComponentBytes),
            })?;
        Ok(self.limits)
    }
}

/// A Numbers resource reported by a semantic-limit error.
///
/// Every resource up to [`Self::FormulaRenderDepth`] can be selected through
/// [`SemanticLimits`]. Formula metadata scanning retains additional fixed
/// adapter-owned ceilings.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
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
    /// UTF-8 bytes retained in names, headers and textual cell values.
    OutputTextBytes,
    /// Formula AST nodes processed across all projected tables.
    FormulaRenderWork,
    /// Nested formula thunk arrays processed by one formula.
    FormulaRenderDepth,
    /// Fixed aggregate bytes of type-6383 category metadata scanned by formula
    /// enrichment.
    FormulaCategoryBytes,
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
            Self::FormulaCategoryBytes => "formula category bytes",
            Self::FormulaWork => "formula work",
            Self::TextBytes => "text bytes",
            Self::FormulaDepth => "formula depth",
        })
    }
}

/// Checked resource ceilings for the semantic projection of one Numbers
/// package or workbook. Every dimension defaults to its hard ceiling.
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
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        objects: MAX_OBJECTS,
        sheets: MAX_SHEETS,
        tables: MAX_TABLES,
        references: MAX_REFERENCES,
        materialized_cells: MAX_MATERIALIZED_CELLS,
        output_text_bytes: MAX_TEXT_BYTES,
        formula_render_work: MAX_FORMULA_RENDER_WORK,
        formula_render_depth: MAX_FORMULA_RENDER_DEPTH,
    };

    /// The default projection profile (every dimension at its hard ceiling).
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// A builder starting from this profile, to tighten or relax some of
    /// its dimensions.
    pub const fn to_builder(self) -> SemanticLimitsBuilder {
        SemanticLimitsBuilder { limits: self }
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

    /// Maximum UTF-8 bytes retained across names, headers and textual cells.
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

    /// Return the configured value of one selectable dimension; fixed
    /// adapter-owned dimensions report `None`.
    #[must_use]
    pub const fn value(self, kind: SemanticLimitKind) -> Option<usize> {
        Some(match kind {
            SemanticLimitKind::Objects => self.objects,
            SemanticLimitKind::Sheets => self.sheets,
            SemanticLimitKind::Tables => self.tables,
            SemanticLimitKind::References => self.references,
            SemanticLimitKind::MaterializedCells => self.materialized_cells,
            SemanticLimitKind::OutputTextBytes => self.output_text_bytes,
            SemanticLimitKind::FormulaRenderWork => self.formula_render_work,
            SemanticLimitKind::FormulaRenderDepth => self.formula_render_depth,
            SemanticLimitKind::FormulaCategoryBytes
            | SemanticLimitKind::FormulaWork
            | SemanticLimitKind::TextBytes
            | SemanticLimitKind::FormulaDepth => return None,
        })
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SemanticLimits`].
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}

fn checked_semantic(kind: SemanticLimitKind, value: usize) -> Result<usize, Error> {
    let maximum = SemanticLimits::MAXIMUM.value(kind).unwrap_or(0);
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind: LimitKind::Semantic(kind),
            value: u64::try_from(value).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    Ok(value)
}

impl SemanticLimitsBuilder {
    /// Set the maximum indexed native objects.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_objects(mut self, value: usize) -> Result<Self, Error> {
        self.limits.objects = checked_semantic(SemanticLimitKind::Objects, value)?;
        Ok(self)
    }

    /// Set the maximum rooted sheets.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_sheets(mut self, value: usize) -> Result<Self, Error> {
        self.limits.sheets = checked_semantic(SemanticLimitKind::Sheets, value)?;
        Ok(self)
    }

    /// Set the maximum semantic tables.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_tables(mut self, value: usize) -> Result<Self, Error> {
        self.limits.tables = checked_semantic(SemanticLimitKind::Tables, value)?;
        Ok(self)
    }

    /// Set the maximum rooted references and retained formula-enrichment
    /// entries.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_references(mut self, value: usize) -> Result<Self, Error> {
        self.limits.references = checked_semantic(SemanticLimitKind::References, value)?;
        Ok(self)
    }

    /// Set the maximum materialized cells across all projected tables.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_materialized_cells(mut self, value: usize) -> Result<Self, Error> {
        self.limits.materialized_cells =
            checked_semantic(SemanticLimitKind::MaterializedCells, value)?;
        Ok(self)
    }

    /// Set the maximum retained UTF-8 bytes across names, headers and
    /// textual cells.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_output_text_bytes(mut self, value: usize) -> Result<Self, Error> {
        self.limits.output_text_bytes =
            checked_semantic(SemanticLimitKind::OutputTextBytes, value)?;
        Ok(self)
    }

    /// Set the maximum formula AST nodes processed across all tables.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_formula_render_work(mut self, value: usize) -> Result<Self, Error> {
        self.limits.formula_render_work =
            checked_semantic(SemanticLimitKind::FormulaRenderWork, value)?;
        Ok(self)
    }

    /// Set the maximum nested thunk depth of one formula.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_formula_render_depth(mut self, value: usize) -> Result<Self, Error> {
        self.limits.formula_render_depth =
            checked_semantic(SemanticLimitKind::FormulaRenderDepth, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Infallible today; the result keeps the builder idiom uniform.
    pub const fn build(self) -> Result<SemanticLimits, Error> {
        Ok(self.limits)
    }
}

/// The dimension named by a Numbers limit error.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// A physical ingress dimension.
    Source(SourceLimitKind),
    /// A semantic projection dimension.
    Semantic(SemanticLimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Source(kind) => kind.fmt(formatter),
            Self::Semantic(kind) => kind.fmt(formatter),
        }
    }
}

impl From<SourceLimitKind> for LimitKind {
    fn from(kind: SourceLimitKind) -> Self {
        Self::Source(kind)
    }
}

impl From<SemanticLimitKind> for LimitKind {
    fn from(kind: SemanticLimitKind) -> Self {
        Self::Semantic(kind)
    }
}

/// Configuration for one Numbers ingress: the physical source profile and
/// the semantic projection profile. Setters consume already validated
/// values and perform no I/O.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct OpenOptions {
    source_limits: SourceLimits,
    semantic_limits: SemanticLimits,
}

impl OpenOptions {
    /// The default ingress policies.
    pub const fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            semantic_limits: SemanticLimits::DEFAULT,
        }
    }

    /// Select the physical source limits.
    pub const fn source_limits(mut self, limits: SourceLimits) -> Self {
        self.source_limits = limits;
        self
    }

    /// Select the semantic projection limits.
    pub const fn semantic_limits(mut self, limits: SemanticLimits) -> Self {
        self.semantic_limits = limits;
        self
    }

    /// The selected physical source limits.
    #[must_use]
    pub const fn source_limits_value(&self) -> SourceLimits {
        self.source_limits
    }

    /// The selected semantic projection limits.
    #[must_use]
    pub const fn semantic_limits_value(&self) -> SemanticLimits {
        self.semantic_limits
    }
}

#[cfg(test)]
mod tests {
    use super::{
        Error, LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind,
        SourceLimits,
    };

    #[test]
    fn builders_reject_zero_and_values_above_the_hard_ceiling() {
        let zero = SourceLimits::builder().max_input_bytes(0).unwrap_err();
        assert!(matches!(
            zero,
            Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                value: 0,
                ..
            }
        ));
        let above = SourceLimits::MAXIMUM.max_members() + 1;
        let error = SourceLimits::builder().max_members(above).unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::Members),
                ..
            }
        ));
        let semantic = SemanticLimits::builder().max_objects(0).unwrap_err();
        assert!(matches!(
            semantic,
            Error::InvalidLimit {
                kind: LimitKind::Semantic(SemanticLimitKind::Objects),
                ..
            }
        ));
        let tables = SemanticLimits::builder()
            .max_tables(SemanticLimits::MAXIMUM.max_tables() + 1)
            .unwrap_err();
        assert!(matches!(
            tables,
            Error::InvalidLimit {
                kind: LimitKind::Semantic(SemanticLimitKind::Tables),
                ..
            }
        ));
    }

    #[test]
    fn selected_dimensions_are_preserved() -> Result<(), Error> {
        let limits = SemanticLimits::builder()
            .max_materialized_cells(17)?
            .max_output_text_bytes(23)?
            .max_formula_render_work(31)?
            .max_formula_render_depth(7)?
            .build()?;
        assert_eq!(limits.max_materialized_cells(), 17);
        assert_eq!(limits.max_output_text_bytes(), 23);
        assert_eq!(limits.max_formula_render_work(), 31);
        assert_eq!(limits.max_formula_render_depth(), 7);
        let source = SourceLimits::builder()
            .max_input_bytes(1024)?
            .max_component_bytes(4096)?
            .build()?;
        assert_eq!(source.max_input_bytes(), 1024);
        assert_eq!(source.max_component_bytes(), 4096);
        Ok(())
    }

    #[test]
    fn default_options_carry_both_default_profiles() {
        let options = OpenOptions::default();
        assert_eq!(options.source_limits_value(), SourceLimits::DEFAULT);
        assert_eq!(options.semantic_limits_value(), SemanticLimits::DEFAULT);
        assert_eq!(
            SourceLimits::DEFAULT.max_input_bytes(),
            litchi_iwa_archive::Limits::MAX_INPUT_BYTES
        );
    }
}
