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

use crate::core::XmlStreamLimits;
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

/// The bounded raw-XML dimension named by an ODF parser failure.
///
/// These dimensions describe parser work and retained namespace state, not
/// semantic document objects.  In particular, namespace scope depth is kept
/// separate from declaration count so a deep document cannot be reported as
/// a declaration overflow.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RawXmlLimitKind {
    /// Total XML input bytes admitted to one raw parser.
    InputBytes,
    /// Element nesting depth observed by the parser.
    Depth,
    /// Number of parser events delivered.
    Events,
    /// Aggregate XML attributes observed by the parser.
    Attributes,
    /// One lexical token's bytes.
    TokenBytes,
    /// Namespace scope frames retained by the resolver.
    NamespaceScopeDepth,
    /// Namespace declarations on one element.
    NamespaceDeclarations,
    /// Active namespace bindings retained across open scopes.
    NamespaceBindings,
    /// Prefix and URI bytes retained by active namespace bindings.
    NamespaceBytes,
    /// One namespace prefix's lexical bytes.
    NamespacePrefixBytes,
    /// One namespace URI's normalized bytes.
    NamespaceUriBytes,
}

impl RawXmlLimitKind {
    /// Every raw XML dimension in stable display order.
    pub const ALL: [Self; 11] = [
        Self::InputBytes,
        Self::Depth,
        Self::Events,
        Self::Attributes,
        Self::TokenBytes,
        Self::NamespaceScopeDepth,
        Self::NamespaceDeclarations,
        Self::NamespaceBindings,
        Self::NamespaceBytes,
        Self::NamespacePrefixBytes,
        Self::NamespaceUriBytes,
    ];

    /// The fixed common-core scope used for typed conversion.
    #[must_use]
    pub const fn scope(self) -> &'static str {
        match self {
            Self::InputBytes => "ODF raw XML input bytes",
            Self::Depth => "ODF raw XML depth",
            Self::Events => "ODF raw XML events",
            Self::Attributes => "ODF raw XML attributes",
            Self::TokenBytes => "ODF raw XML token bytes",
            Self::NamespaceScopeDepth => "ODF raw XML namespace scope depth",
            Self::NamespaceDeclarations => "ODF raw XML namespace declarations",
            Self::NamespaceBindings => "ODF raw XML namespace bindings",
            Self::NamespaceBytes => "ODF raw XML namespace bytes",
            Self::NamespacePrefixBytes => "ODF raw XML namespace prefix bytes",
            Self::NamespaceUriBytes => "ODF raw XML namespace URI bytes",
        }
    }

    /// Classify a common-core resource-limit scope without string inference.
    pub(crate) fn from_scope(scope: &str) -> Option<Self> {
        Some(match scope {
            "ODF raw XML input bytes" => Self::InputBytes,
            "ODF raw XML depth" => Self::Depth,
            "ODF raw XML events" => Self::Events,
            "ODF raw XML attributes" => Self::Attributes,
            "ODF raw XML token bytes" => Self::TokenBytes,
            "ODF raw XML namespace scope depth" => Self::NamespaceScopeDepth,
            "ODF raw XML namespace declarations" => Self::NamespaceDeclarations,
            "ODF raw XML namespace bindings" => Self::NamespaceBindings,
            "ODF raw XML namespace bytes" => Self::NamespaceBytes,
            "ODF raw XML namespace prefix bytes" => Self::NamespacePrefixBytes,
            "ODF raw XML namespace URI bytes" => Self::NamespaceUriBytes,
            _ => return None,
        })
    }
}

impl fmt::Display for RawXmlLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let name = match self {
            Self::InputBytes => "raw_xml_input_bytes",
            Self::Depth => "raw_xml_depth",
            Self::Events => "raw_xml_events",
            Self::Attributes => "raw_xml_attributes",
            Self::TokenBytes => "raw_xml_token_bytes",
            Self::NamespaceScopeDepth => "raw_xml_namespace_scope_depth",
            Self::NamespaceDeclarations => "raw_xml_namespace_declarations",
            Self::NamespaceBindings => "raw_xml_namespace_bindings",
            Self::NamespaceBytes => "raw_xml_namespace_bytes",
            Self::NamespacePrefixBytes => "raw_xml_namespace_prefix_bytes",
            Self::NamespaceUriBytes => "raw_xml_namespace_uri_bytes",
        };
        formatter.write_str(name)
    }
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
    /// Bytes of the decoded `content.xml` part, and of the XML of a flat (single-file) document.
    ContentBytes,
    /// Bytes of the decoded `styles.xml` part.
    StylesBytes,
    /// Top-level semantic units of the projection: paragraphs and tables of
    /// a text document, worksheets of a workbook, slides of a presentation,
    /// pages of a drawing.
    Units,
    /// XML element nesting depth inspected by a flat-document or chart
    /// reader.
    Depth,
    /// UTF-8 bytes of one staged replacement text (a flat paragraph, shape
    /// or slide text, an image-map area text).
    TextBytes,
    /// Staged edits of one flat-document edit.
    Edits,
    /// Drawing shapes of a flat drawing.
    Shapes,
    /// Image frames of a flat image document.
    Frames,
    /// Image-map areas of a flat image document.
    MapAreas,
    /// Bytes of one published flat document.
    OutputBytes,
    /// Chart axes.
    Axes,
    /// Chart series.
    Series,
    /// Chart data points.
    DataPoints,
    /// Rows of a chart's cached data table.
    CachedRows,
    /// Cells of a chart's cached data table.
    CachedCells,
    /// Items of one chart cell-range address list.
    RangeItems,
    /// Package resources a chart references.
    Resources,
    /// Retained undo/redo steps of a chart history.
    History,
    /// Bytes of one chart scalar (a name, label or formula).
    ScalarBytes,
}

impl SemanticLimitKind {
    /// Every semantic dimension, in declaration order.
    pub const ALL: [Self; 19] = [
        Self::ContentBytes,
        Self::StylesBytes,
        Self::Units,
        Self::Depth,
        Self::TextBytes,
        Self::Edits,
        Self::Shapes,
        Self::Frames,
        Self::MapAreas,
        Self::OutputBytes,
        Self::Axes,
        Self::Series,
        Self::DataPoints,
        Self::CachedRows,
        Self::CachedCells,
        Self::RangeItems,
        Self::Resources,
        Self::History,
        Self::ScalarBytes,
    ];

    /// The registered glossary noun of this dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ContentBytes => "content_bytes",
            Self::StylesBytes => "styles_bytes",
            Self::Units => "units",
            Self::Depth => "depth",
            Self::TextBytes => "text_bytes",
            Self::Edits => "edits",
            Self::Shapes => "shapes",
            Self::Frames => "frames",
            Self::MapAreas => "map_areas",
            Self::OutputBytes => "output_bytes",
            Self::Axes => "axes",
            Self::Series => "series",
            Self::DataPoints => "data_points",
            Self::CachedRows => "cached_rows",
            Self::CachedCells => "cached_cells",
            Self::RangeItems => "range_items",
            Self::Resources => "resources",
            Self::History => "history",
            Self::ScalarBytes => "scalar_bytes",
        }
    }

    /// The common-core resource-limit scope naming this dimension.
    pub(crate) const fn scope(self) -> &'static str {
        match self {
            Self::ContentBytes => "ODF semantic content_bytes",
            Self::StylesBytes => "ODF semantic styles_bytes",
            Self::Units => "ODF semantic units",
            Self::Depth => "ODF semantic depth",
            Self::TextBytes => "ODF semantic text_bytes",
            Self::Edits => "ODF semantic edits",
            Self::Shapes => "ODF semantic shapes",
            Self::Frames => "ODF semantic frames",
            Self::MapAreas => "ODF semantic map_areas",
            Self::OutputBytes => "ODF semantic output_bytes",
            Self::Axes => "ODF semantic axes",
            Self::Series => "ODF semantic series",
            Self::DataPoints => "ODF semantic data_points",
            Self::CachedRows => "ODF semantic cached_rows",
            Self::CachedCells => "ODF semantic cached_cells",
            Self::RangeItems => "ODF semantic range_items",
            Self::Resources => "ODF semantic resources",
            Self::History => "ODF semantic history",
            Self::ScalarBytes => "ODF semantic scalar_bytes",
        }
    }

    /// Classify a common-core resource-limit scope raised for a semantic
    /// dimension.
    pub(crate) fn from_scope(scope: &str) -> Option<Self> {
        Self::ALL.into_iter().find(|kind| kind.scope() == scope)
    }

    /// The common-core resource class of this dimension.
    #[must_use]
    pub(crate) const fn resource(self) -> litchi_core::Resource {
        match self {
            Self::ContentBytes | Self::StylesBytes | Self::TextBytes | Self::ScalarBytes => {
                litchi_core::Resource::InputBytes
            },
            Self::Units
            | Self::Edits
            | Self::Shapes
            | Self::Frames
            | Self::MapAreas
            | Self::Axes
            | Self::Series
            | Self::DataPoints
            | Self::CachedRows
            | Self::CachedCells
            | Self::RangeItems
            | Self::Resources
            | Self::History => litchi_core::Resource::Objects,
            Self::Depth => litchi_core::Resource::Depth,
            Self::OutputBytes => litchi_core::Resource::OutputBytes,
        }
    }
}

impl fmt::Display for SemanticLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Checked limits applied when an ODF package is projected into its semantic
/// model (`Document`, `Workbook` or `Presentation`), when a flat
/// (single-file) document is opened or edited, and when a chart is read.
///
/// The decoded part dimensions are checked against the declared member sizes
/// before the part is inflated, and `units` is checked once the projection's
/// top-level units have been counted. A flat document's XML is bounded by
/// `content_bytes` (and by the source `input_bytes`); its nesting, staged
/// edits and replacement text by `depth`, `edits` and `text_bytes`. The chart
/// dimensions bound the ODC chart model. Each default is the most permissive
/// bound an ODF reader enforced before the dimension became configurable, so
/// a default open accepts everything it accepted before.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct SemanticLimits {
    content_bytes: u64,
    styles_bytes: u64,
    units: u64,
    depth: u64,
    text_bytes: u64,
    edits: u64,
    shapes: u64,
    frames: u64,
    map_areas: u64,
    output_bytes: u64,
    axes: u64,
    series: u64,
    data_points: u64,
    cached_rows: u64,
    cached_cells: u64,
    range_items: u64,
    resources: u64,
    history: u64,
    scalar_bytes: u64,
}

impl SemanticLimits {
    /// The hard ceilings: no builder accepts a larger value.
    pub const MAXIMUM: Self = Self {
        content_bytes: 256 * MIB,
        styles_bytes: 256 * MIB,
        units: 1_048_576,
        depth: 4_096,
        text_bytes: 16 * MIB,
        edits: 1_024,
        shapes: 1_000_000,
        frames: 1_000_000,
        map_areas: 100_000,
        output_bytes: 256 * MIB,
        axes: 16_777_216,
        series: 16_777_216,
        data_points: 16_777_216,
        cached_rows: 16_777_216,
        cached_cells: 16_777_216,
        range_items: 16_777_216,
        resources: 16_777_216,
        history: 16_777_216,
        scalar_bytes: 16 * MIB,
    };

    /// The default profile.
    pub const DEFAULT: Self = Self {
        content_bytes: 256 * MIB,
        styles_bytes: 256 * MIB,
        units: 1_048_576,
        depth: 4_096,
        text_bytes: 16 * MIB,
        edits: 1_024,
        shapes: 1_000_000,
        frames: 1_000_000,
        map_areas: 100_000,
        output_bytes: 256 * MIB,
        axes: 16_384,
        series: 65_536,
        data_points: 16_777_216,
        cached_rows: 16_777_216,
        cached_cells: 16_777_216,
        range_items: 65_536,
        resources: 100_000,
        history: 4_096,
        scalar_bytes: 64 * 1024,
    };

    /// Start a checked builder from [`Self::DEFAULT`].
    pub const fn builder() -> SemanticLimitsBuilder {
        SemanticLimitsBuilder {
            limits: Self::DEFAULT,
        }
    }

    /// The selected maximum of [`SemanticLimitKind::ContentBytes`].
    #[must_use]
    pub const fn max_content_bytes(self) -> u64 {
        self.content_bytes
    }

    /// The selected maximum of [`SemanticLimitKind::StylesBytes`].
    #[must_use]
    pub const fn max_styles_bytes(self) -> u64 {
        self.styles_bytes
    }

    /// The selected maximum of [`SemanticLimitKind::Units`].
    #[must_use]
    pub const fn max_units(self) -> u64 {
        self.units
    }

    /// The selected maximum of [`SemanticLimitKind::Depth`].
    #[must_use]
    pub const fn max_depth(self) -> u64 {
        self.depth
    }

    /// The selected maximum of [`SemanticLimitKind::TextBytes`].
    #[must_use]
    pub const fn max_text_bytes(self) -> u64 {
        self.text_bytes
    }

    /// The selected maximum of [`SemanticLimitKind::Edits`].
    #[must_use]
    pub const fn max_edits(self) -> u64 {
        self.edits
    }

    /// The selected maximum of [`SemanticLimitKind::Shapes`].
    #[must_use]
    pub const fn max_shapes(self) -> u64 {
        self.shapes
    }

    /// The selected maximum of [`SemanticLimitKind::Frames`].
    #[must_use]
    pub const fn max_frames(self) -> u64 {
        self.frames
    }

    /// The selected maximum of [`SemanticLimitKind::MapAreas`].
    #[must_use]
    pub const fn max_map_areas(self) -> u64 {
        self.map_areas
    }

    /// The selected maximum of [`SemanticLimitKind::OutputBytes`].
    #[must_use]
    pub const fn max_output_bytes(self) -> u64 {
        self.output_bytes
    }

    /// The selected maximum of [`SemanticLimitKind::Axes`].
    #[must_use]
    pub const fn max_axes(self) -> u64 {
        self.axes
    }

    /// The selected maximum of [`SemanticLimitKind::Series`].
    #[must_use]
    pub const fn max_series(self) -> u64 {
        self.series
    }

    /// The selected maximum of [`SemanticLimitKind::DataPoints`].
    #[must_use]
    pub const fn max_data_points(self) -> u64 {
        self.data_points
    }

    /// The selected maximum of [`SemanticLimitKind::CachedRows`].
    #[must_use]
    pub const fn max_cached_rows(self) -> u64 {
        self.cached_rows
    }

    /// The selected maximum of [`SemanticLimitKind::CachedCells`].
    #[must_use]
    pub const fn max_cached_cells(self) -> u64 {
        self.cached_cells
    }

    /// The selected maximum of [`SemanticLimitKind::RangeItems`].
    #[must_use]
    pub const fn max_range_items(self) -> u64 {
        self.range_items
    }

    /// The selected maximum of [`SemanticLimitKind::Resources`].
    #[must_use]
    pub const fn max_resources(self) -> u64 {
        self.resources
    }

    /// The selected maximum of [`SemanticLimitKind::History`].
    #[must_use]
    pub const fn max_history(self) -> u64 {
        self.history
    }

    /// The selected maximum of [`SemanticLimitKind::ScalarBytes`].
    #[must_use]
    pub const fn max_scalar_bytes(self) -> u64 {
        self.scalar_bytes
    }

    /// Return the configured value of one dimension.
    #[must_use]
    pub const fn value(self, kind: SemanticLimitKind) -> u64 {
        match kind {
            SemanticLimitKind::ContentBytes => self.content_bytes,
            SemanticLimitKind::StylesBytes => self.styles_bytes,
            SemanticLimitKind::Units => self.units,
            SemanticLimitKind::Depth => self.depth,
            SemanticLimitKind::TextBytes => self.text_bytes,
            SemanticLimitKind::Edits => self.edits,
            SemanticLimitKind::Shapes => self.shapes,
            SemanticLimitKind::Frames => self.frames,
            SemanticLimitKind::MapAreas => self.map_areas,
            SemanticLimitKind::OutputBytes => self.output_bytes,
            SemanticLimitKind::Axes => self.axes,
            SemanticLimitKind::Series => self.series,
            SemanticLimitKind::DataPoints => self.data_points,
            SemanticLimitKind::CachedRows => self.cached_rows,
            SemanticLimitKind::CachedCells => self.cached_cells,
            SemanticLimitKind::RangeItems => self.range_items,
            SemanticLimitKind::Resources => self.resources,
            SemanticLimitKind::History => self.history,
            SemanticLimitKind::ScalarBytes => self.scalar_bytes,
        }
    }

    /// Whether this profile is the default (no caller tightening).
    #[must_use]
    pub const fn is_default(self) -> bool {
        let mut index = 0;
        while index < SemanticLimitKind::ALL.len() {
            let kind = SemanticLimitKind::ALL[index];
            if self.value(kind) != Self::DEFAULT.value(kind) {
                return false;
            }
            index += 1;
        }
        true
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

    /// Check one observed `usize` count or size against this profile.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Limit`] when `observed` exceeds the dimension.
    pub fn check_usize(self, kind: SemanticLimitKind, observed: usize) -> Result<(), crate::Error> {
        self.check(kind, u64::try_from(observed).unwrap_or(u64::MAX))
    }

    /// Check one observed value against this profile for a reader that
    /// returns the common-core result: an exceeded dimension is a
    /// [`litchi_core::Error::ResourceLimit`] whose scope names the dimension,
    /// which converting into [`crate::Error`] types back into
    /// [`crate::Error::Limit`].
    ///
    /// # Errors
    ///
    /// Returns the resource-limit failure when `observed` exceeds the
    /// dimension.
    pub fn check_core(self, kind: SemanticLimitKind, observed: usize) -> litchi_core::Result<()> {
        let maximum = self.value(kind);
        let observed = u64::try_from(observed).unwrap_or(u64::MAX);
        if observed > maximum {
            return Err(litchi_core::Error::ResourceLimit(
                litchi_core::ResourceLimit {
                    resource: kind.resource(),
                    observed,
                    limit: maximum,
                    scope: kind.scope().into(),
                },
            ));
        }
        Ok(())
    }

    /// The selected maximum of one dimension as a `usize` (saturating on
    /// targets whose `usize` is narrower than the value).
    #[must_use]
    pub fn usize_value(self, kind: SemanticLimitKind) -> usize {
        usize::try_from(self.value(kind)).unwrap_or(usize::MAX)
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
    /// Set the maximum of [`SemanticLimitKind::ContentBytes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_content_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.content_bytes = checked_semantic(SemanticLimitKind::ContentBytes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::StylesBytes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_styles_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.styles_bytes = checked_semantic(SemanticLimitKind::StylesBytes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Units`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_units(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.units = checked_semantic(SemanticLimitKind::Units, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Depth`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_depth(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.depth = checked_semantic(SemanticLimitKind::Depth, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::TextBytes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_text_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.text_bytes = checked_semantic(SemanticLimitKind::TextBytes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Edits`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_edits(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.edits = checked_semantic(SemanticLimitKind::Edits, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Shapes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_shapes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.shapes = checked_semantic(SemanticLimitKind::Shapes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Frames`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_frames(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.frames = checked_semantic(SemanticLimitKind::Frames, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::MapAreas`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_map_areas(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.map_areas = checked_semantic(SemanticLimitKind::MapAreas, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::OutputBytes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_output_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.output_bytes = checked_semantic(SemanticLimitKind::OutputBytes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Axes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_axes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.axes = checked_semantic(SemanticLimitKind::Axes, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Series`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_series(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.series = checked_semantic(SemanticLimitKind::Series, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::DataPoints`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_data_points(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.data_points = checked_semantic(SemanticLimitKind::DataPoints, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::CachedRows`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_cached_rows(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.cached_rows = checked_semantic(SemanticLimitKind::CachedRows, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::CachedCells`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_cached_cells(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.cached_cells = checked_semantic(SemanticLimitKind::CachedCells, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::RangeItems`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_range_items(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.range_items = checked_semantic(SemanticLimitKind::RangeItems, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::Resources`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_resources(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.resources = checked_semantic(SemanticLimitKind::Resources, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::History`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_history(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.history = checked_semantic(SemanticLimitKind::History, value)?;
        Ok(self)
    }

    /// Set the maximum of [`SemanticLimitKind::ScalarBytes`].
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::InvalidLimit`] outside `1..=MAXIMUM`.
    pub fn max_scalar_bytes(mut self, value: u64) -> Result<Self, crate::Error> {
        self.limits.scalar_bytes = checked_semantic(SemanticLimitKind::ScalarBytes, value)?;
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
    /// A raw XML parser dimension.
    RawXml(RawXmlLimitKind),
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Source(kind) => kind.fmt(formatter),
            Self::Semantic(kind) => kind.fmt(formatter),
            Self::RawXml(kind) => kind.fmt(formatter),
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

impl From<RawXmlLimitKind> for LimitKind {
    fn from(kind: RawXmlLimitKind) -> Self {
        Self::RawXml(kind)
    }
}

/// Configuration for one packaged ODF ingress.
///
/// Setters consume already validated policy values and perform no I/O.
/// Passwords are supplied to the password-bearing routes and never retained
/// here.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct OpenOptions {
    source_limits: SourceLimits,
    semantic_limits: SemanticLimits,
    xml_limits: XmlStreamLimits,
}

impl Default for OpenOptions {
    fn default() -> Self {
        Self::new()
    }
}

impl OpenOptions {
    /// The default ingress policies.
    pub const fn new() -> Self {
        Self {
            source_limits: SourceLimits::DEFAULT,
            semantic_limits: SemanticLimits::DEFAULT,
            // Packaged owners historically admit their complete bounded XML
            // member.  Keep the direct raw scanner's narrow `Default`
            // separate; an explicit `.xml_limits(...)` always replaces this
            // owner ingress profile exactly.
            xml_limits: XmlStreamLimits::owner_default_value(),
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

    /// Select the bounded raw XML parser policy retained by an opened owner.
    ///
    /// The policy is independent of physical package admission and semantic
    /// projection limits.  Keeping it in the ingress envelope ensures that
    /// a later projection or edit does not silently fall back to a parser
    /// default.
    pub const fn xml_limits(mut self, limits: XmlStreamLimits) -> Self {
        self.xml_limits = limits;
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

    /// The selected bounded raw XML parser policy.
    #[must_use]
    pub const fn xml_limits_value(&self) -> XmlStreamLimits {
        self.xml_limits
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
