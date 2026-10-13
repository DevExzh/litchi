//! Caller-selected ingress limits and the ODF open-options envelope.
//!
//! Every packaged ODF owner (`odt`, `ods`, `odp` and the smaller families)
//! accepts one [`OpenOptions`] value on its configured ingress routes.  The
//! envelope carries checked physical [`SourceLimits`] (input bytes, ZIP
//! catalog and mandatory member ceilings) and checked [`SemanticLimits`]
//! (projection dimensions enforced after the package has been parsed).
//!
//! Construction follows the workspace idiom: a consuming builder with fallible
//! `max_<dimension>` setters and a fallible `build()`.  A value outside
//! `1..=maximum` is reported as [`crate::Error::InvalidLimit`]; an exceeded
//! dimension at ingress is reported as [`crate::Error::Limit`].

use std::fmt;

use crate::package::ArchiveLimits;

const MIB: u64 = 1024 * 1024;

/// The physical ingress dimension named by a source limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Bytes of the complete physical source (file, buffer or stream).
    InputBytes,
    /// Non-directory ZIP members in the central directory.
    Members,
    /// Raw bytes of one ZIP member name.
    ArchiveMemberNameBytes,
    /// Aggregate central-directory metadata bytes.
    ArchiveMetadataBytes,
    /// Declared compressed bytes of one ZIP member.
    ArchiveCompressedBytes,
    /// Declared uncompressed bytes of one ZIP member.
    EntryBytes,
    /// Aggregate declared uncompressed bytes of all ZIP members.
    AggregateBytes,
    /// Bytes of the mandatory `mimetype` member.
    MimetypeBytes,
    /// Bytes of the `META-INF/manifest.xml` member.
    ManifestBytes,
}

impl SourceLimitKind {
    /// Every source dimension, in declaration order.
    pub const ALL: [Self; 9] = [
        Self::InputBytes,
        Self::Members,
        Self::ArchiveMemberNameBytes,
        Self::ArchiveMetadataBytes,
        Self::ArchiveCompressedBytes,
        Self::EntryBytes,
        Self::AggregateBytes,
        Self::MimetypeBytes,
        Self::ManifestBytes,
    ];

    /// The registered glossary noun of this dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::InputBytes => "input_bytes",
            Self::Members => "members",
            Self::ArchiveMemberNameBytes => "archive_member_name_bytes",
            Self::ArchiveMetadataBytes => "archive_metadata_bytes",
            Self::ArchiveCompressedBytes => "archive_compressed_bytes",
            Self::EntryBytes => "entry_bytes",
            Self::AggregateBytes => "aggregate_bytes",
            Self::MimetypeBytes => "mimetype_bytes",
            Self::ManifestBytes => "manifest_bytes",
        }
    }

    /// Classify a common-core resource-limit scope raised by the ODF package
    /// owner.  Scopes are structural constants, never authored content.
    pub(crate) fn from_scope(scope: &str) -> Option<Self> {
        Some(match scope {
            "ODF positional package" | "ODF owned package input" => Self::InputBytes,
            "ODF ZIP file count" => Self::Members,
            "ODF ZIP member name bytes" => Self::ArchiveMemberNameBytes,
            "ODF ZIP metadata bytes" => Self::ArchiveMetadataBytes,
            "ODF ZIP compressed bytes" => Self::ArchiveCompressedBytes,
            "ODF ZIP entry bytes" => Self::EntryBytes,
            "ODF ZIP total bytes" => Self::AggregateBytes,
            "ODF mimetype bytes" => Self::MimetypeBytes,
            "ODF manifest bytes" => Self::ManifestBytes,
            _ => return None,
        })
    }

    /// The common-core scope used when this kind is flattened into a
    /// [`litchi_core::ResourceLimit`].
    pub(crate) const fn scope(self) -> &'static str {
        match self {
            Self::InputBytes => "ODF owned package input",
            Self::Members => "ODF ZIP file count",
            Self::ArchiveMemberNameBytes => "ODF ZIP member name bytes",
            Self::ArchiveMetadataBytes => "ODF ZIP metadata bytes",
            Self::ArchiveCompressedBytes => "ODF ZIP compressed bytes",
            Self::EntryBytes => "ODF ZIP entry bytes",
            Self::AggregateBytes => "ODF ZIP total bytes",
            Self::MimetypeBytes => "ODF mimetype bytes",
            Self::ManifestBytes => "ODF manifest bytes",
        }
    }
}

impl fmt::Display for SourceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Checked physical limits for one ODF package ingress.
///
/// `max_input_bytes` bounds the physical source before ZIP indexing starts;
/// the archive dimensions bound the central directory and the declared member
/// sizes; the mandatory-member dimensions bound the `mimetype` and manifest
/// payloads read before any other member.  Every dimension is finite.  The
/// defaults use the workspace-shared 512 MiB input ceiling; callers may
/// tighten any dimension or raise one up to [`SourceLimits::MAXIMUM`].
///
/// ```
/// use litchi_odf_common::{OpenOptions, SourceLimits};
///
/// # fn main() -> Result<(), litchi_odf_common::Error> {
/// let limits = SourceLimits::builder()
///     .max_input_bytes(32 * 1024 * 1024)?
///     .max_members(4_096)?
///     .build()?;
/// let options = OpenOptions::new().source_limits(limits);
/// assert_eq!(options.source_limits_value().max_members(), 4_096);
/// # Ok(())
/// # }
/// ```
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: u64,
    archive: ArchiveLimits,
    mimetype_bytes: u64,
    manifest_bytes: u64,
}

impl SourceLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        input_bytes: 2 * 1024 * MIB,
        archive: ArchiveLimits::DEFAULT,
        mimetype_bytes: 4 * 1024,
        manifest_bytes: 16 * MIB,
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

    /// Maximum bytes of the complete physical source.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }

    /// Maximum non-directory ZIP members.
    #[must_use]
    pub const fn max_members(self) -> usize {
        self.archive.max_files()
    }

    /// Maximum raw bytes of one ZIP member name.
    #[must_use]
    pub const fn max_archive_member_name_bytes(self) -> u64 {
        self.archive.max_member_name_bytes()
    }

    /// Maximum aggregate central-directory metadata bytes.
    #[must_use]
    pub const fn max_archive_metadata_bytes(self) -> u64 {
        self.archive.max_metadata_bytes()
    }

    /// Maximum declared compressed bytes of one ZIP member.
    #[must_use]
    pub const fn max_archive_compressed_bytes(self) -> u64 {
        self.archive.max_compressed_size()
    }

    /// Maximum declared uncompressed bytes of one ZIP member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.archive.max_entry_size()
    }

    /// Maximum aggregate declared uncompressed bytes.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.archive.max_total_size()
    }

    /// Maximum bytes of the `mimetype` member.
    #[must_use]
    pub const fn max_mimetype_bytes(self) -> u64 {
        self.mimetype_bytes
    }

    /// Maximum bytes of the manifest member.
    #[must_use]
    pub const fn max_manifest_bytes(self) -> u64 {
        self.manifest_bytes
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SourceLimitKind) -> u64 {
        match kind {
            SourceLimitKind::InputBytes => self.input_bytes,
            SourceLimitKind::Members => self.archive.max_files() as u64,
            SourceLimitKind::ArchiveMemberNameBytes => self.archive.max_member_name_bytes(),
            SourceLimitKind::ArchiveMetadataBytes => self.archive.max_metadata_bytes(),
            SourceLimitKind::ArchiveCompressedBytes => self.archive.max_compressed_size(),
            SourceLimitKind::EntryBytes => self.archive.max_entry_size(),
            SourceLimitKind::AggregateBytes => self.archive.max_total_size(),
            SourceLimitKind::MimetypeBytes => self.mimetype_bytes,
            SourceLimitKind::ManifestBytes => self.manifest_bytes,
        }
    }

    /// The ZIP catalog profile handed to the archive indexer.
    pub(crate) const fn archive(self) -> ArchiveLimits {
        self.archive
    }

    /// Derive a profile from a detector input ceiling and catalog profile.
    /// Values are clamped to the hard ceilings, never rejected: the detector
    /// validated its own profile before calling this.
    pub(crate) fn from_detection(input_bytes: u64, archive: ArchiveLimits) -> Self {
        let maximum = Self::MAXIMUM.archive;
        Self {
            input_bytes: input_bytes.clamp(1, Self::MAXIMUM.input_bytes),
            archive: ArchiveLimits::new(
                archive.max_files().clamp(1, maximum.max_files()),
                archive
                    .max_member_name_bytes()
                    .clamp(1, maximum.max_member_name_bytes()),
                archive
                    .max_metadata_bytes()
                    .clamp(1, maximum.max_metadata_bytes()),
                archive
                    .max_compressed_size()
                    .clamp(1, maximum.max_compressed_size()),
                archive.max_entry_size().clamp(1, maximum.max_entry_size()),
                archive.max_total_size().clamp(1, maximum.max_total_size()),
            ),
            ..Self::MAXIMUM
        }
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SourceLimits`].
#[derive(Clone, Copy, Debug)]
#[must_use]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

fn checked(kind: SourceLimitKind, value: u64) -> Result<u64, crate::Error> {
    let maximum = SourceLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(crate::Error::InvalidLimit {
            kind: LimitKind::Source(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

impl SourceLimitsBuilder {
    /// Set the maximum physical source bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_input_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.input_bytes = checked(SourceLimitKind::InputBytes, value)?;
        Ok(self)
    }

    /// Set the maximum non-directory ZIP members.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_members(mut self, value: usize) -> Result<Self, crate::Error> {
        let value = checked(
            SourceLimitKind::Members,
            u64::try_from(value).unwrap_or(u64::MAX),
        )?;
        // The maximum is a `usize` value, so a checked value fits.
        self.limits.archive = self
            .limits
            .archive
            .with_max_files(usize::try_from(value).unwrap_or(usize::MAX));
        Ok(self)
    }

    /// Set the maximum raw bytes of one ZIP member name.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_archive_member_name_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        let value = checked(SourceLimitKind::ArchiveMemberNameBytes, value)?;
        self.limits.archive = self.limits.archive.with_max_member_name_bytes(value);
        Ok(self)
    }

    /// Set the maximum aggregate central-directory metadata bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_archive_metadata_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        let value = checked(SourceLimitKind::ArchiveMetadataBytes, value)?;
        self.limits.archive = self.limits.archive.with_max_metadata_bytes(value);
        Ok(self)
    }

    /// Set the maximum declared compressed bytes of one ZIP member.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_archive_compressed_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        let value = checked(SourceLimitKind::ArchiveCompressedBytes, value)?;
        self.limits.archive = self.limits.archive.with_max_compressed_size(value);
        Ok(self)
    }

    /// Set the maximum declared uncompressed bytes of one ZIP member.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_entry_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        let value = checked(SourceLimitKind::EntryBytes, value)?;
        self.limits.archive = self.limits.archive.with_max_entry_size(value);
        Ok(self)
    }

    /// Set the maximum aggregate declared uncompressed bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_aggregate_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        let value = checked(SourceLimitKind::AggregateBytes, value)?;
        self.limits.archive = self.limits.archive.with_max_total_size(value);
        Ok(self)
    }

    /// Set the maximum bytes of the `mimetype` member.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_mimetype_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.mimetype_bytes = checked(SourceLimitKind::MimetypeBytes, value)?;
        Ok(self)
    }

    /// Set the maximum bytes of the manifest member.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_manifest_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.manifest_bytes = checked(SourceLimitKind::ManifestBytes, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Every setter already validated its dimension; this re-validates the
    /// complete profile so a future cross-dimension rule has one home.
    pub fn build(self) -> Result<SourceLimits, crate::Error> {
        for kind in SourceLimitKind::ALL {
            checked(kind, self.limits.value(kind))?;
        }
        Ok(self.limits)
    }
}

/// The semantic projection dimension named by a semantic limit.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SemanticLimitKind {
    /// Bytes of the decoded `content.xml` part.
    ContentBytes,
    /// Bytes of the decoded `styles.xml` part.
    StylesBytes,
    /// Top-level semantic units of the projection: paragraphs and tables of
    /// a text document, worksheets of a workbook, slides of a presentation.
    Units,
}

impl SemanticLimitKind {
    /// Every semantic dimension, in declaration order.
    pub const ALL: [Self; 3] = [Self::ContentBytes, Self::StylesBytes, Self::Units];

    /// The registered glossary noun of this dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ContentBytes => "content_bytes",
            Self::StylesBytes => "styles_bytes",
            Self::Units => "units",
        }
    }
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Checked limits applied when an ODF package is projected into its semantic
/// model (`Document`, `Workbook` or `Presentation`).
///
/// The decoded part dimensions are checked against the declared member sizes
/// before the part is inflated, and `units` is checked once the projection's
/// top-level units have been counted.  The defaults equal the hard ceilings
/// the readers already enforced (256 MiB parts, 1,048,576 units), so a
/// default open behaves exactly as before.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SemanticLimits {
    content_bytes: u64,
    styles_bytes: u64,
    units: u64,
}

impl SemanticLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        content_bytes: 256 * MIB,
        styles_bytes: 256 * MIB,
        units: 1_048_576,
    };

    /// The default profile, equal to [`Self::MAXIMUM`].
    pub const DEFAULT: Self = Self::MAXIMUM;

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// Maximum decoded `content.xml` bytes.
    #[must_use]
    pub const fn max_content_bytes(self) -> u64 {
        self.content_bytes
    }

    /// Maximum decoded `styles.xml` bytes.
    #[must_use]
    pub const fn max_styles_bytes(self) -> u64 {
        self.styles_bytes
    }

    /// Maximum top-level semantic units of the projection.
    #[must_use]
    pub const fn max_units(self) -> u64 {
        self.units
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SemanticLimitKind) -> u64 {
        match kind {
            SemanticLimitKind::ContentBytes => self.content_bytes,
            SemanticLimitKind::StylesBytes => self.styles_bytes,
            SemanticLimitKind::Units => self.units,
        }
    }

    /// Whether this profile is the default (no caller tightening).
    #[must_use]
    pub const fn is_default(self) -> bool {
        self.content_bytes == Self::DEFAULT.content_bytes
            && self.styles_bytes == Self::DEFAULT.styles_bytes
            && self.units == Self::DEFAULT.units
    }

    /// Check one observed value against this profile.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Limit`] when `observed` exceeds the dimension.
    pub fn check(self, kind: SemanticLimitKind, observed: u64) -> Result<(), crate::Error> {
        let maximum = self.value(kind);
        if observed > maximum {
            return Err(crate::Error::Limit {
                kind: LimitKind::Semantic(kind),
                observed,
                maximum,
            });
        }
        Ok(())
    }
}

impl Default for SemanticLimits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// Consuming builder for [`SemanticLimits`].
#[derive(Clone, Copy, Debug)]
#[must_use]
pub struct SemanticLimitsBuilder {
    limits: SemanticLimits,
}

fn checked_semantic(kind: SemanticLimitKind, value: u64) -> Result<u64, crate::Error> {
    let maximum = SemanticLimits::MAXIMUM.value(kind);
    if value == 0 || value > maximum {
        return Err(crate::Error::InvalidLimit {
            kind: LimitKind::Semantic(kind),
            value,
            maximum,
        });
    }
    Ok(value)
}

impl SemanticLimitsBuilder {
    /// Set the maximum decoded `content.xml` bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_content_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.content_bytes = checked_semantic(SemanticLimitKind::ContentBytes, value)?;
        Ok(self)
    }

    /// Set the maximum decoded `styles.xml` bytes.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_styles_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.styles_bytes = checked_semantic(SemanticLimitKind::StylesBytes, value)?;
        Ok(self)
    }

    /// Set the maximum top-level semantic units.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_units(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.units = checked_semantic(SemanticLimitKind::Units, value)?;
        Ok(self)
    }

    /// Finish the profile.
    ///
    /// # Errors
    ///
    /// Re-validates every dimension.
    pub fn build(self) -> Result<SemanticLimits, crate::Error> {
        for kind in SemanticLimitKind::ALL {
            checked_semantic(kind, self.limits.value(kind))?;
        }
        Ok(self.limits)
    }
}

/// The dimension named by an ODF limit error.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
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

/// Configuration for one packaged ODF ingress.
///
/// Setters consume already validated policy values and perform no I/O.
/// Passwords are supplied to the password-bearing routes and never retained
/// here.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
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
        LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits, SourceLimitKind, SourceLimits,
    };
    use crate::Error;

    #[test]
    fn builder_rejects_zero_and_values_above_the_hard_ceiling() {
        let zero = SourceLimits::builder().max_input_bytes(0).unwrap_err();
        assert!(matches!(
            zero,
            Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::InputBytes),
                value: 0,
                ..
            }
        ));
        let above = SourceLimits::MAXIMUM.max_manifest_bytes() + 1;
        let error = SourceLimits::builder()
            .max_manifest_bytes(above)
            .unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidLimit {
                kind: LimitKind::Source(SourceLimitKind::ManifestBytes),
                value,
                maximum,
            } if value == above && maximum == above - 1
        ));
        let semantic = SemanticLimits::builder().max_units(0).unwrap_err();
        assert!(matches!(
            semantic,
            Error::InvalidLimit {
                kind: LimitKind::Semantic(SemanticLimitKind::Units),
                ..
            }
        ));
    }

    #[test]
    fn defaults_share_the_workspace_input_ceiling() {
        assert_eq!(SourceLimits::default().max_input_bytes(), 512 * 1024 * 1024);
        assert_eq!(OpenOptions::new(), OpenOptions::default());
        let limits = SourceLimits::builder()
            .max_input_bytes(SourceLimits::MAXIMUM.max_input_bytes())
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(limits.max_input_bytes(), 2 * 1024 * 1024 * 1024);
        for kind in SourceLimitKind::ALL {
            assert_eq!(SourceLimitKind::from_scope(kind.scope()), Some(kind));
        }
    }
}
