//! Checked ingress limits and the open configuration envelope (ADR 0034
//! §5).
//!
//! [`SourceLimits`] bounds the physical package (input bytes, ZIP members,
//! member and aggregate decoded bytes, one IWA component stream);
//! [`SemanticLimits`] bounds the semantic projection (sections, retained
//! text). Both are built with consuming builders whose `max_<dimension>`
//! setters reject zero and values above the hard ceiling with
//! [`PackageError::InvalidLimit`]. [`OpenOptions`] carries both profiles into
//! [`Package::open_with_options`](crate::Package::open_with_options) and the
//! byte routes.
//!
//! ```
//! use litchi_pages::{OpenOptions, SemanticLimits, SourceLimits};
//!
//! # fn main() -> Result<(), litchi_pages::PackageError> {
//! let options = OpenOptions::new()
//!     .source_limits(SourceLimits::builder().max_input_bytes(64 * 1024 * 1024)?.build()?)
//!     .semantic_limits(SemanticLimits::builder().max_sections(64)?.build()?);
//! assert_eq!(options.source_limits_value().max_input_bytes(), 64 * 1024 * 1024);
//! assert_eq!(options.semantic_limits_value().max_sections(), 64);
//! # Ok(())
//! # }
//! ```

use std::fmt;

use crate::PackageError;
use crate::document::{DEFAULT_MAX_TEXT_BYTES, MAX_SECTIONS};

const MIB: u64 = 1024 * 1024;

/// A physical ingress dimension of [`SourceLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Encoded bytes of the whole source (file, byte slice or directory).
    InputBytes,
    /// ZIP members (or directory files) admitted from the source.
    Members,
    /// Decoded bytes of one member.
    EntryBytes,
    /// Aggregate decoded bytes of every member.
    AggregateBytes,
    /// Decompressed bytes of one IWA component stream.
    ComponentBytes,
}

impl SourceLimitKind {
    /// Every source dimension, in declaration order.
    pub const ALL: [Self; 5] = [
        Self::InputBytes,
        Self::Members,
        Self::EntryBytes,
        Self::AggregateBytes,
        Self::ComponentBytes,
    ];
}

impl fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::Members => "members",
            Self::EntryBytes => "entry bytes",
            Self::AggregateBytes => "aggregate bytes",
            Self::ComponentBytes => "component bytes",
        })
    }
}

/// A semantic projection dimension of [`SemanticLimits`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Ordered sections decoded from the Pages body graph.
    Sections,
    /// Aggregate UTF-8 bytes retained by semantic section names and text.
    TextBytes,
}

impl SemanticLimitKind {
    /// Every semantic dimension, in declaration order.
    pub const ALL: [Self; 2] = [Self::Sections, Self::TextBytes];
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::Sections => "sections",
            Self::TextBytes => "text bytes",
        })
    }
}

/// The dimension named by a limit error.
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

/// Checked physical limits for one Pages package ingress.
///
/// The defaults use the workspace-shared 512 MiB input ceiling; every other
/// dimension starts at its hard ceiling ([`Self::MAXIMUM`]). A builder may
/// tighten any dimension, or raise one up to its ceiling.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: u64,
    members: usize,
    entry_bytes: u64,
    aggregate_bytes: u64,
    component_bytes: usize,
}

impl SourceLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        input_bytes: litchi_iwa_archive::Limits::MAX_INPUT_BYTES,
        members: litchi_iwa_archive::Limits::MAX_ENTRIES,
        entry_bytes: litchi_iwa_archive::Limits::MAX_ENTRY_BYTES,
        aggregate_bytes: litchi_iwa_archive::Limits::MAX_TOTAL_BYTES,
        component_bytes: litchi_iwa_archive::Limits::MAX_IWA_STREAM_BYTES,
    };

    /// The default ingress profile (512 MiB input; other dimensions at their
    /// hard ceilings).
    pub const DEFAULT: Self = Self {
        input_bytes: 512 * MIB,
        ..Self::MAXIMUM
    };

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum encoded source bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum ZIP members (or directory files).
    #[must_use]
    pub const fn max_members(self) -> usize {
        self.members
    }

    /// Maximum decoded bytes of one member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }

    /// Maximum aggregate decoded bytes.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }

    /// Maximum decompressed bytes of one IWA component stream.
    #[must_use]
    pub const fn max_component_bytes(self) -> usize {
        self.component_bytes
    }

    const fn value(self, kind: SourceLimitKind) -> u64 {
        match kind {
            SourceLimitKind::InputBytes => self.input_bytes,
            SourceLimitKind::Members => self.members as u64,
            SourceLimitKind::EntryBytes => self.entry_bytes,
            SourceLimitKind::AggregateBytes => self.aggregate_bytes,
            SourceLimitKind::ComponentBytes => self.component_bytes as u64,
        }
    }

    /// The shared iWork archive profile with these dimensions.
    pub(crate) fn archive(self) -> Result<litchi_iwa_archive::Limits, PackageError> {
        Ok(litchi_iwa_archive::Limits::new(
            self.input_bytes,
            self.members,
            self.entry_bytes,
            self.aggregate_bytes,
            self.component_bytes,
        )?)
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn checked_source(kind: SourceLimitKind, value: u64) -> Result<u64, PackageError> {
    let maximum = SourceLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(PackageError::InvalidLimit {
            kind: LimitKind::Source(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

fn usize_value(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

/// Consuming builder for [`SourceLimits`].
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Set the maximum encoded source bytes.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_input_bytes(mut self, value: u64) -> Result<Self, PackageError> {
        self.limits.input_bytes = checked_source(SourceLimitKind::InputBytes, value)?;
        Ok(self)
    }

    /// Set the maximum ZIP members (or directory files).
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_members(mut self, value: usize) -> Result<Self, PackageError> {
        checked_source(SourceLimitKind::Members, usize_value(value))?;
        self.limits.members = value;
        Ok(self)
    }

    /// Set the maximum decoded bytes of one member.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_entry_bytes(mut self, value: u64) -> Result<Self, PackageError> {
        self.limits.entry_bytes = checked_source(SourceLimitKind::EntryBytes, value)?;
        Ok(self)
    }

    /// Set the maximum aggregate decoded bytes.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_aggregate_bytes(mut self, value: u64) -> Result<Self, PackageError> {
        self.limits.aggregate_bytes = checked_source(SourceLimitKind::AggregateBytes, value)?;
        Ok(self)
    }

    /// Set the maximum decompressed bytes of one IWA component stream.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_component_bytes(mut self, value: usize) -> Result<Self, PackageError> {
        checked_source(SourceLimitKind::ComponentBytes, usize_value(value))?;
        self.limits.component_bytes = value;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Re-validates every dimension.
    pub fn build(self) -> Result<SourceLimits, PackageError> {
        for kind in SourceLimitKind::ALL {
            checked_source(kind, self.limits.value(kind))?;
        }
        Ok(self.limits)
    }
}

/// Checked semantic projection limits for one Pages document.
///
/// The defaults are the hard ceilings (4096 sections, 64 MiB retained text);
/// a builder may only tighten them.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SemanticLimits {
    sections: usize,
    text_bytes: usize,
}

impl SemanticLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        sections: MAX_SECTIONS,
        text_bytes: DEFAULT_MAX_TEXT_BYTES,
    };

    /// The default projection profile (the hard ceilings).
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum ordered sections admitted by the projection.
    #[must_use]
    pub const fn max_sections(self) -> usize {
        self.sections
    }

    /// Maximum aggregate UTF-8 bytes retained by the projection.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.text_bytes
    }

    const fn value(self, kind: SemanticLimitKind) -> u64 {
        match kind {
            SemanticLimitKind::Sections => self.sections as u64,
            SemanticLimitKind::TextBytes => self.text_bytes as u64,
        }
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

fn checked_semantic(kind: SemanticLimitKind, value: u64) -> Result<u64, PackageError> {
    let maximum = SemanticLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(PackageError::InvalidLimit {
            kind: LimitKind::Semantic(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

/// Consuming builder for [`SemanticLimits`].
#[derive(Debug, Clone, Copy)]
#[must_use]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}

impl SemanticLimitsBuilder {
    /// Set the maximum ordered sections.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_sections(mut self, value: usize) -> Result<Self, PackageError> {
        checked_semantic(SemanticLimitKind::Sections, usize_value(value))?;
        self.limits.sections = value;
        Ok(self)
    }

    /// Set the maximum aggregate retained text bytes.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_text_bytes(mut self, value: usize) -> Result<Self, PackageError> {
        checked_semantic(SemanticLimitKind::TextBytes, usize_value(value))?;
        self.limits.text_bytes = value;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Re-validates every dimension.
    pub fn build(self) -> Result<SemanticLimits, PackageError> {
        for kind in SemanticLimitKind::ALL {
            checked_semantic(kind, self.limits.value(kind))?;
        }
        Ok(self.limits)
    }
}

/// Configuration for one Pages package ingress.
///
/// Setters consume already validated profiles and perform no I/O.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[must_use]
pub struct OpenOptions {
    source_limits: SourceLimits,
    semantic_limits: SemanticLimits,
    archive_limits: Option<litchi_iwa_archive::Limits>,
}

impl OpenOptions {
    /// The default ingress policies.
    pub const fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            semantic_limits: SemanticLimits::DEFAULT,
            archive_limits: None,
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

    /// Select the complete shared iWork archive profile directly, overriding
    /// the archive profile derived from [`Self::source_limits`].
    ///
    /// Unstable seam for the iWork crates' own tests and migration hosts,
    /// which tune archive-internal dimensions (message bytes, header fields)
    /// that [`SourceLimits`] does not name.
    #[doc(hidden)]
    pub const fn __archive_limits(mut self, limits: litchi_iwa_archive::Limits) -> Self {
        self.archive_limits = Some(limits);
        self
    }

    /// The archive profile this envelope selects.
    pub(crate) fn archive(&self) -> Result<litchi_iwa_archive::Limits, PackageError> {
        match self.archive_limits {
            Some(limits) => Ok(limits),
            None => self.source_limits.archive(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::{
        LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
    };
    use crate::PackageError;

    #[test]
    fn builders_reject_zero_and_values_above_the_ceiling() {
        let zero = SourceLimits::builder().max_input_bytes(0).unwrap_err();
        assert!(matches!(
            zero,
            PackageError::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                value: 0,
                ..
            }
        ));
        let above = SourceLimits::MAXIMUM.max_members() + 1;
        let error = SourceLimits::builder().max_members(above).unwrap_err();
        assert!(matches!(
            error,
            PackageError::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::Members),
                ..
            }
        ));
        let semantic = SemanticLimits::builder()
            .max_sections(SemanticLimits::MAXIMUM.max_sections() + 1)
            .unwrap_err();
        assert!(matches!(
            semantic,
            PackageError::InvalidLimit {
                kind: LimitKind::Semantic(SemanticLimitKind::Sections),
                ..
            }
        ));
    }

    #[test]
    fn defaults_share_the_workspace_input_ceiling() {
        assert_eq!(SourceLimits::default().max_input_bytes(), 512 * 1024 * 1024);
        assert_eq!(SemanticLimits::default(), SemanticLimits::MAXIMUM);
        let options = OpenOptions::default();
        assert_eq!(options.source_limits_value(), SourceLimits::DEFAULT);
        assert!(SourceLimits::DEFAULT.archive().is_ok());
    }
}
