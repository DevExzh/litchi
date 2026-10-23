//! Checked resource policy for reading OPC packages.
//!
//! The policy is deliberately owned by the OPC boundary. Format crates can
//! forward one profile without copying ZIP, XML, relationship, or graph
//! ceilings into their own public APIs.

use crate::{OpcError, Result};
use soapberry_zip::office::ArchiveLimits;
use std::fmt;

const KIB: usize = 1024;
const MIB: usize = 1024 * KIB;
const GIB: u64 = 1024 * 1024 * 1024;

/// A bounded OPC resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SourceLimitKind {
    /// Compressed package input bytes.
    InputBytes,
    /// ZIP non-directory members.
    Members,
    /// Total ZIP central-directory entries, including directory records.
    ArchiveTotalEntries,
    /// ZIP member name bytes.
    ArchiveMemberNameBytes,
    /// ZIP central-directory metadata bytes.
    ArchiveMetadataBytes,
    /// ZIP member compressed bytes.
    ArchiveCompressedBytes,
    /// ZIP member declared uncompressed bytes.
    EntryBytes,
    /// ZIP aggregate declared uncompressed bytes.
    AggregateBytes,
    /// OPC parts materialized from the archive.
    Parts,
    /// One materialized OPC part's bytes.
    PartBytes,
    /// Aggregate materialized OPC part bytes.
    TotalPartBytes,
    /// `[Content_Types].xml` bytes.
    ContentTypesBytes,
    /// Content-type default and override mappings.
    ContentTypeMappings,
    /// Relationship parts.
    RelationshipParts,
    /// One relationships XML part's bytes.
    RelationshipXmlBytes,
    /// Aggregate relationships XML bytes.
    TotalRelationshipXmlBytes,
    /// Relationships in one relationships part.
    RelationshipsPerPart,
    /// Aggregate relationships in a package.
    TotalRelationships,
    /// Nodes discovered while traversing relationship targets.
    RelationshipGraphNodes,
    /// XML events in one checked XML part or operation.
    Events,
    /// Aggregate XML events across relationship parts.
    TotalRelationshipXmlEvents,
    /// XML nesting depth.
    Depth,
    /// One XML attribute's bytes.
    XmlAttributeBytes,
    /// One relationship target's bytes.
    RelationshipTargetBytes,
}

impl fmt::Display for SourceLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "package input bytes",
            Self::Members => "ZIP members",
            Self::ArchiveTotalEntries => "ZIP total entries",
            Self::ArchiveMemberNameBytes => "ZIP member name bytes",
            Self::ArchiveMetadataBytes => "ZIP metadata bytes",
            Self::ArchiveCompressedBytes => "ZIP member compressed bytes",
            Self::EntryBytes => "ZIP member uncompressed bytes",
            Self::AggregateBytes => "ZIP aggregate uncompressed bytes",
            Self::Parts => "OPC parts",
            Self::PartBytes => "OPC part bytes",
            Self::TotalPartBytes => "aggregate OPC part bytes",
            Self::ContentTypesBytes => "content-types XML bytes",
            Self::ContentTypeMappings => "content-type mappings",
            Self::RelationshipParts => "relationship parts",
            Self::RelationshipXmlBytes => "relationship XML bytes",
            Self::TotalRelationshipXmlBytes => "aggregate relationship XML bytes",
            Self::RelationshipsPerPart => "relationships per part",
            Self::TotalRelationships => "aggregate relationships",
            Self::RelationshipGraphNodes => "relationship graph nodes",
            Self::Events => "XML events",
            Self::TotalRelationshipXmlEvents => "aggregate relationship XML events",
            Self::Depth => "XML nesting depth",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::RelationshipTargetBytes => "relationship target bytes",
        })
    }
}

/// Checked resource ceilings for reading one OPC package.
///
/// Create a profile with [`SourceLimits::builder`]. Every public ingress takes
/// the selected profile, so callers can choose tighter bounds for hostile or
/// multi-tenant input without an unbounded escape hatch.
///
/// The default profile is intentionally bounded across compressed input, ZIP
/// structure and declared sizes, materialized parts, content types, and
/// relationship parsing and traversal. Its ceilings are a Litchi safety policy
/// rather than maxima defined by ECMA-376 Part 2 sections 7.3.6 and 10 or
/// MS-OI29500 sections 2.1.1749-1752.
#[allow(
    clippy::module_name_repetitions,
    reason = "physical ingress profile follows the shared SourceLimits vocabulary"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceLimits {
    input_bytes: u64,
    members: usize,
    archive_total_entries: usize,
    archive_member_name_bytes: u64,
    archive_metadata_bytes: u64,
    archive_compressed_bytes: u64,
    entry_bytes: u64,
    aggregate_bytes: u64,
    parts: usize,
    part_bytes: u64,
    total_part_bytes: u64,
    content_types_bytes: usize,
    content_type_mappings: usize,
    relationship_parts: usize,
    relationship_xml_bytes: usize,
    total_relationship_xml_bytes: usize,
    relationships_per_part: usize,
    total_relationships: usize,
    relationship_graph_nodes: usize,
    events: usize,
    total_relationship_xml_events: usize,
    depth: usize,
    xml_attribute_bytes: usize,
    relationship_target_bytes: usize,
}

impl SourceLimits {
    /// Start building a checked read policy from the standard defaults.
    #[must_use]
    pub fn builder() -> SourceLimitsBuilder {
        SourceLimitsBuilder {
            limits: Self::default(),
        }
    }

    /// Intersect two source policies component by component.
    ///
    /// A format operation may tighten the policy supplied at open time, but
    /// it must never widen that retained policy.  The componentwise minimum
    /// preserves the consistency relationships checked by
    /// [`SourceLimitsBuilder::build`].
    #[must_use]
    pub fn min(self, other: Self) -> Self {
        Self {
            input_bytes: self.input_bytes.min(other.input_bytes),
            members: self.members.min(other.members),
            archive_total_entries: self.archive_total_entries.min(other.archive_total_entries),
            archive_member_name_bytes: self
                .archive_member_name_bytes
                .min(other.archive_member_name_bytes),
            archive_metadata_bytes: self
                .archive_metadata_bytes
                .min(other.archive_metadata_bytes),
            archive_compressed_bytes: self
                .archive_compressed_bytes
                .min(other.archive_compressed_bytes),
            entry_bytes: self.entry_bytes.min(other.entry_bytes),
            aggregate_bytes: self.aggregate_bytes.min(other.aggregate_bytes),
            parts: self.parts.min(other.parts),
            part_bytes: self.part_bytes.min(other.part_bytes),
            total_part_bytes: self.total_part_bytes.min(other.total_part_bytes),
            content_types_bytes: self.content_types_bytes.min(other.content_types_bytes),
            content_type_mappings: self.content_type_mappings.min(other.content_type_mappings),
            relationship_parts: self.relationship_parts.min(other.relationship_parts),
            relationship_xml_bytes: self
                .relationship_xml_bytes
                .min(other.relationship_xml_bytes),
            total_relationship_xml_bytes: self
                .total_relationship_xml_bytes
                .min(other.total_relationship_xml_bytes),
            relationships_per_part: self
                .relationships_per_part
                .min(other.relationships_per_part),
            total_relationships: self.total_relationships.min(other.total_relationships),
            relationship_graph_nodes: self
                .relationship_graph_nodes
                .min(other.relationship_graph_nodes),
            events: self.events.min(other.events),
            total_relationship_xml_events: self
                .total_relationship_xml_events
                .min(other.total_relationship_xml_events),
            depth: self.depth.min(other.depth),
            xml_attribute_bytes: self.xml_attribute_bytes.min(other.xml_attribute_bytes),
            relationship_target_bytes: self
                .relationship_target_bytes
                .min(other.relationship_target_bytes),
        }
    }

    /// Maximum accepted compressed package input bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }
    /// Maximum non-directory ZIP members.
    #[must_use]
    pub const fn max_members(self) -> usize {
        self.members
    }
    /// Maximum total ZIP central-directory entries, including directories.
    #[must_use]
    pub const fn max_archive_total_entries(self) -> usize {
        self.archive_total_entries
    }
    /// Maximum bytes in one ZIP member name.
    #[must_use]
    pub const fn max_archive_member_name_bytes(self) -> u64 {
        self.archive_member_name_bytes
    }
    /// Maximum ZIP central-directory metadata bytes.
    #[must_use]
    pub const fn max_archive_metadata_bytes(self) -> u64 {
        self.archive_metadata_bytes
    }
    /// Maximum compressed bytes declared for one ZIP member.
    #[must_use]
    pub const fn max_archive_compressed_bytes(self) -> u64 {
        self.archive_compressed_bytes
    }
    /// Maximum declared uncompressed bytes for one ZIP member.
    #[must_use]
    pub const fn max_entry_bytes(self) -> u64 {
        self.entry_bytes
    }
    /// Maximum aggregate declared uncompressed ZIP bytes.
    #[must_use]
    pub const fn max_aggregate_bytes(self) -> u64 {
        self.aggregate_bytes
    }
    /// Maximum materialized OPC parts.
    #[must_use]
    pub const fn max_parts(self) -> usize {
        self.parts
    }
    /// Maximum bytes in one materialized OPC part.
    #[must_use]
    pub const fn max_part_bytes(self) -> u64 {
        self.part_bytes
    }
    /// Maximum aggregate bytes across materialized OPC parts.
    #[must_use]
    pub const fn max_total_part_bytes(self) -> u64 {
        self.total_part_bytes
    }
    /// Maximum `[Content_Types].xml` bytes.
    #[must_use]
    pub const fn max_content_types_bytes(self) -> usize {
        self.content_types_bytes
    }
    /// Maximum content-type mappings.
    #[must_use]
    pub const fn max_content_type_mappings(self) -> usize {
        self.content_type_mappings
    }
    /// Maximum relationship parts.
    #[must_use]
    pub const fn max_relationship_parts(self) -> usize {
        self.relationship_parts
    }
    /// Maximum bytes in one relationship XML part.
    #[must_use]
    pub const fn max_relationship_xml_bytes(self) -> usize {
        self.relationship_xml_bytes
    }
    /// Maximum aggregate relationship XML bytes.
    #[must_use]
    pub const fn max_total_relationship_xml_bytes(self) -> usize {
        self.total_relationship_xml_bytes
    }
    /// Maximum relationships in one relationship part.
    #[must_use]
    pub const fn max_relationships_per_part(self) -> usize {
        self.relationships_per_part
    }
    /// Maximum relationships across a package.
    #[must_use]
    pub const fn max_total_relationships(self) -> usize {
        self.total_relationships
    }
    /// Maximum relationship graph nodes.
    #[must_use]
    pub const fn max_relationship_graph_nodes(self) -> usize {
        self.relationship_graph_nodes
    }
    /// Maximum XML events in one checked XML part or operation.
    #[must_use]
    pub const fn max_events(self) -> usize {
        self.events
    }
    /// Maximum XML events across relationship parts.
    #[must_use]
    pub const fn max_total_relationship_xml_events(self) -> usize {
        self.total_relationship_xml_events
    }
    /// Maximum nesting depth in checked XML.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.depth
    }
    /// Maximum bytes in any XML attribute.
    #[must_use]
    pub const fn max_xml_attribute_bytes(self) -> usize {
        self.xml_attribute_bytes
    }
    /// Maximum bytes in a relationship target.
    #[must_use]
    pub const fn max_relationship_target_bytes(self) -> usize {
        self.relationship_target_bytes
    }

    /// Convert this profile to the ZIP indexing policy.
    #[must_use]
    pub(crate) const fn zip_limits(self) -> ArchiveLimits {
        ArchiveLimits {
            max_files: self.members,
            max_member_name_bytes: self.archive_member_name_bytes,
            max_metadata_bytes: self.archive_metadata_bytes,
            max_compressed_size: self.archive_compressed_bytes,
            max_entry_size: self.entry_bytes,
            max_total_size: self.aggregate_bytes,
        }
    }

    #[allow(
        clippy::unused_self,
        reason = "kept as a method so crate-internal callers write limits.check(...)"
    )]
    pub(crate) fn check(self, resource: SourceLimitKind, actual: u64, maximum: u64) -> Result<()> {
        if actual > maximum {
            return Err(OpcError::Limit {
                kind: resource,
                observed: actual,
                maximum,
            });
        }
        Ok(())
    }

    pub(crate) fn check_input_bytes(self, actual: u64) -> Result<()> {
        self.check(SourceLimitKind::InputBytes, actual, self.input_bytes)
    }
}

impl Default for SourceLimits {
    fn default() -> Self {
        Self {
            input_bytes: 512 * MIB as u64,
            members: 100_000,
            archive_total_entries: 100_000,
            archive_member_name_bytes: 4 * KIB as u64,
            archive_metadata_bytes: 64 * MIB as u64,
            archive_compressed_bytes: 512 * MIB as u64,
            entry_bytes: 512 * MIB as u64,
            aggregate_bytes: 2 * GIB,
            parts: 100_000,
            part_bytes: 512 * MIB as u64,
            total_part_bytes: 512 * MIB as u64,
            content_types_bytes: 8 * MIB,
            content_type_mappings: 100_000,
            relationship_parts: 100_000,
            relationship_xml_bytes: 8 * MIB,
            total_relationship_xml_bytes: 64 * MIB,
            relationships_per_part: 100_000,
            total_relationships: 1_000_000,
            relationship_graph_nodes: 100_000,
            events: 1_000_000,
            total_relationship_xml_events: 8_000_000,
            depth: 256,
            xml_attribute_bytes: 64 * KIB,
            relationship_target_bytes: 4 * KIB,
        }
    }
}

/// Builder for [`SourceLimits`].
#[derive(Debug, Clone, Copy)]
pub struct SourceLimitsBuilder {
    limits: SourceLimits,
}

impl SourceLimitsBuilder {
    /// Finalize this checked profile.
    ///
    /// # Errors
    ///
    /// Returns `Err` if the ceilings are inconsistent: relationship parts
    /// exceed archive members, parts exceed archive members, or relationship
    /// target bytes exceed XML attribute bytes.
    pub fn build(self) -> Result<SourceLimits> {
        if self.limits.relationship_parts > self.limits.members {
            return Err(invalid(
                SourceLimitKind::RelationshipParts,
                self.limits.relationship_parts as u64,
                self.limits.members as u64,
            ));
        }
        if self.limits.members > self.limits.archive_total_entries {
            return Err(invalid(
                SourceLimitKind::Members,
                self.limits.members as u64,
                self.limits.archive_total_entries as u64,
            ));
        }
        if self.limits.parts > self.limits.members {
            return Err(invalid(
                SourceLimitKind::Parts,
                self.limits.parts as u64,
                self.limits.members as u64,
            ));
        }
        if self.limits.relationship_target_bytes > self.limits.xml_attribute_bytes {
            return Err(invalid(
                SourceLimitKind::RelationshipTargetBytes,
                self.limits.relationship_target_bytes as u64,
                self.limits.xml_attribute_bytes as u64,
            ));
        }
        Ok(self.limits)
    }

    /// Set the compressed input ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero or exceeds the addressable range.
    pub fn max_input_bytes(mut self, value: u64) -> Result<Self> {
        validate_input(value)?;
        self.limits.input_bytes = value;
        Ok(self)
    }
    /// Set the ZIP member count ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_members(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::Members, value)?;
        self.limits.members = value;
        Ok(self)
    }
    /// Set the total ZIP central-directory entry ceiling, including directories.
    pub fn max_archive_total_entries(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::ArchiveTotalEntries, value)?;
        self.limits.archive_total_entries = value;
        Ok(self)
    }
    /// Set the ZIP member-name byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_member_name_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::ArchiveMemberNameBytes, value)?;
        self.limits.archive_member_name_bytes = value;
        Ok(self)
    }
    /// Set the ZIP metadata byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_metadata_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::ArchiveMetadataBytes, value)?;
        self.limits.archive_metadata_bytes = value;
        Ok(self)
    }
    /// Set the per-member ZIP compressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_compressed_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::ArchiveCompressedBytes, value)?;
        self.limits.archive_compressed_bytes = value;
        Ok(self)
    }
    /// Set the per-member ZIP uncompressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_entry_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::EntryBytes, value)?;
        self.limits.entry_bytes = value;
        Ok(self)
    }
    /// Set the aggregate ZIP uncompressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_aggregate_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::AggregateBytes, value)?;
        self.limits.aggregate_bytes = value;
        Ok(self)
    }
    /// Set the materialized OPC part count ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_parts(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::Parts, value)?;
        self.limits.parts = value;
        Ok(self)
    }
    /// Set the per-part byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_part_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::PartBytes, value)?;
        self.limits.part_bytes = value;
        Ok(self)
    }
    /// Set the aggregate materialized-part byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_part_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(SourceLimitKind::TotalPartBytes, value)?;
        self.limits.total_part_bytes = value;
        Ok(self)
    }
    /// Set the content-types XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_content_types_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::ContentTypesBytes, value)?;
        self.limits.content_types_bytes = value;
        Ok(self)
    }
    /// Set the content-type mapping ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_content_type_mappings(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::ContentTypeMappings, value)?;
        self.limits.content_type_mappings = value;
        Ok(self)
    }
    /// Set the relationship-part ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_parts(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::RelationshipParts, value)?;
        self.limits.relationship_parts = value;
        Ok(self)
    }
    /// Set the per-relationship-part XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_xml_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::RelationshipXmlBytes, value)?;
        self.limits.relationship_xml_bytes = value;
        Ok(self)
    }
    /// Set the aggregate relationship XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationship_xml_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::TotalRelationshipXmlBytes, value)?;
        self.limits.total_relationship_xml_bytes = value;
        Ok(self)
    }
    /// Set the relationships-per-part ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationships_per_part(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::RelationshipsPerPart, value)?;
        self.limits.relationships_per_part = value;
        Ok(self)
    }
    /// Set the aggregate relationship ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationships(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::TotalRelationships, value)?;
        self.limits.total_relationships = value;
        Ok(self)
    }
    /// Set the relationship graph-node ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_graph_nodes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::RelationshipGraphNodes, value)?;
        self.limits.relationship_graph_nodes = value;
        Ok(self)
    }
    /// Set the per-part XML parser event ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_events(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::Events, value)?;
        self.limits.events = value;
        Ok(self)
    }
    /// Set the aggregate relationship XML event ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationship_xml_events(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::TotalRelationshipXmlEvents, value)?;
        self.limits.total_relationship_xml_events = value;
        Ok(self)
    }
    /// Set the XML nesting ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_depth(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::Depth, value)?;
        self.limits.depth = value;
        Ok(self)
    }
    /// Set the universal XML attribute byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_xml_attribute_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::XmlAttributeBytes, value)?;
        self.limits.xml_attribute_bytes = value;
        Ok(self)
    }
    /// Set the relationship target byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_target_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(SourceLimitKind::RelationshipTargetBytes, value)?;
        self.limits.relationship_target_bytes = value;
        Ok(self)
    }
}

fn invalid(kind: SourceLimitKind, value: u64, maximum: u64) -> OpcError {
    OpcError::InvalidLimit {
        kind,
        value,
        maximum,
    }
}

fn validate_input(value: u64) -> Result<()> {
    if value == 0 || value > usize::MAX.saturating_sub(1) as u64 {
        return Err(invalid(
            SourceLimitKind::InputBytes,
            value,
            usize::MAX.saturating_sub(1) as u64,
        ));
    }
    Ok(())
}

fn validate_bytes(resource: SourceLimitKind, value: u64) -> Result<()> {
    if value == 0 {
        return Err(invalid(resource, value, u64::MAX));
    }
    Ok(())
}

fn validate_count(resource: SourceLimitKind, value: usize) -> Result<()> {
    if value == 0 {
        return Err(invalid(resource, 0, usize::MAX as u64));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic"
    )]
    use super::*;

    #[test]
    fn builder_rejects_zero_and_keeps_defaults_checked() {
        assert!(matches!(
            SourceLimits::builder().max_input_bytes(0),
            Err(OpcError::InvalidLimit {
                kind: SourceLimitKind::InputBytes,
                value: 0,
                ..
            })
        ));
        assert_eq!(
            SourceLimits::builder()
                .build()
                .unwrap()
                .max_total_part_bytes(),
            512 * MIB as u64
        );
    }

    fn invalid_fields(error: OpcError) -> (SourceLimitKind, u64, u64) {
        match error {
            OpcError::InvalidLimit {
                kind,
                value,
                maximum,
            } => (kind, value, maximum),
            error => panic!("expected invalid limit, got {error:?}"),
        }
    }

    #[test]
    fn invalid_scalar_limits_report_their_actual_upper_bounds() {
        let input_max = usize::MAX.saturating_sub(1) as u64;
        for value in [0, u64::MAX] {
            assert_eq!(
                invalid_fields(SourceLimits::builder().max_input_bytes(value).unwrap_err()),
                (SourceLimitKind::InputBytes, value, input_max)
            );
        }
        assert_eq!(
            invalid_fields(SourceLimits::builder().max_entry_bytes(0).unwrap_err()),
            (SourceLimitKind::EntryBytes, 0, u64::MAX)
        );
        assert_eq!(
            invalid_fields(SourceLimits::builder().max_depth(0).unwrap_err()),
            (SourceLimitKind::Depth, 0, usize::MAX as u64)
        );
        assert!(SourceLimits::builder().max_input_bytes(input_max).is_ok());
        assert!(SourceLimits::builder().max_entry_bytes(u64::MAX).is_ok());
        assert!(SourceLimits::builder().max_depth(usize::MAX).is_ok());
        assert_eq!(
            invalid_fields(
                crate::SourcePartSpliceLimits::new(u64::MAX, 1, 1, 1, 1, 1).unwrap_err()
            ),
            (SourceLimitKind::PartBytes, u64::MAX, u64::MAX - 1)
        );
    }

    #[test]
    fn inconsistent_profiles_report_the_offending_dimension_and_bound() {
        let cases = [
            (
                SourceLimits::builder()
                    .max_members(1)
                    .unwrap()
                    .max_relationship_parts(2)
                    .unwrap()
                    .build()
                    .unwrap_err(),
                (SourceLimitKind::RelationshipParts, 2, 1),
            ),
            (
                SourceLimits::builder()
                    .max_members(3)
                    .unwrap()
                    .max_relationship_parts(1)
                    .unwrap()
                    .max_parts(1)
                    .unwrap()
                    .max_archive_total_entries(2)
                    .unwrap()
                    .build()
                    .unwrap_err(),
                (SourceLimitKind::Members, 3, 2),
            ),
            (
                SourceLimits::builder()
                    .max_members(1)
                    .unwrap()
                    .max_relationship_parts(1)
                    .unwrap()
                    .max_parts(2)
                    .unwrap()
                    .build()
                    .unwrap_err(),
                (SourceLimitKind::Parts, 2, 1),
            ),
            (
                SourceLimits::builder()
                    .max_xml_attribute_bytes(2)
                    .unwrap()
                    .max_relationship_target_bytes(3)
                    .unwrap()
                    .build()
                    .unwrap_err(),
                (SourceLimitKind::RelationshipTargetBytes, 3, 2),
            ),
        ];
        for (error, expected) in cases {
            assert_eq!(invalid_fields(error), expected);
        }
    }

    #[test]
    fn source_limit_diagnostics_keep_canonical_content_free_fields() {
        let error = OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            observed: 2,
            maximum: 1,
        };
        assert_eq!(
            format!("{error:?}"),
            "Limit { kind: InputBytes, observed: 2, maximum: 1 }"
        );
        assert_eq!(
            format!("{:?}", crate::error::replicate_deferred_error(&error)),
            format!("{error:?}")
        );
        let invalid = SourceLimits::builder().max_input_bytes(0).unwrap_err();
        assert_eq!(
            format!("{invalid:?}"),
            format!(
                "InvalidLimit {{ kind: InputBytes, value: 0, maximum: {} }}",
                usize::MAX.saturating_sub(1)
            )
        );
    }
}
