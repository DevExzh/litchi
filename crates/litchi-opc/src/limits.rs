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
pub enum ReadResource {
    /// Compressed package input bytes.
    InputBytes,
    /// ZIP non-directory members.
    ArchiveMembers,
    /// ZIP member name bytes.
    ArchiveMemberNameBytes,
    /// ZIP central-directory metadata bytes.
    ArchiveMetadataBytes,
    /// ZIP member compressed bytes.
    ArchiveCompressedBytes,
    /// ZIP member declared uncompressed bytes.
    ArchiveEntryBytes,
    /// ZIP aggregate declared uncompressed bytes.
    ArchiveTotalBytes,
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
    /// XML events in one relationships part.
    XmlEvents,
    /// Aggregate XML events across relationship parts.
    TotalRelationshipXmlEvents,
    /// XML nesting depth.
    XmlDepth,
    /// One XML attribute's bytes.
    XmlAttributeBytes,
    /// One relationship target's bytes.
    RelationshipTargetBytes,
}

impl fmt::Display for ReadResource {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "package input bytes",
            Self::ArchiveMembers => "ZIP members",
            Self::ArchiveMemberNameBytes => "ZIP member name bytes",
            Self::ArchiveMetadataBytes => "ZIP metadata bytes",
            Self::ArchiveCompressedBytes => "ZIP member compressed bytes",
            Self::ArchiveEntryBytes => "ZIP member uncompressed bytes",
            Self::ArchiveTotalBytes => "ZIP aggregate uncompressed bytes",
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
            Self::XmlEvents => "relationship XML events",
            Self::TotalRelationshipXmlEvents => "aggregate relationship XML events",
            Self::XmlDepth => "XML nesting depth",
            Self::XmlAttributeBytes => "XML attribute bytes",
            Self::RelationshipTargetBytes => "relationship target bytes",
        })
    }
}

/// Checked resource ceilings for reading one OPC package.
///
/// Create a profile with [`ReadLimits::builder`]. Existing open methods use
/// [`ReadLimits::default`], while `*_with_limits` constructors let callers
/// select tighter bounds for hostile or multi-tenant input.
///
/// The default profile is intentionally bounded across compressed input, ZIP
/// structure and declared sizes, materialized parts, content types, and
/// relationship parsing and traversal. Its ceilings are a Litchi safety policy
/// rather than maxima defined by ECMA-376 Part 2 sections 7.3.6 and 10 or
/// MS-OI29500 sections 2.1.1749-1752.
#[allow(
    clippy::module_name_repetitions,
    reason = "profile type is named for its read-policy domain, mirroring ReadResource"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ReadLimits {
    input_bytes: u64,
    archive_members: usize,
    archive_member_name_bytes: u64,
    archive_metadata_bytes: u64,
    archive_compressed_bytes: u64,
    archive_entry_bytes: u64,
    archive_total_bytes: u64,
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
    xml_events: usize,
    total_relationship_xml_events: usize,
    xml_depth: usize,
    xml_attribute_bytes: usize,
    relationship_target_bytes: usize,
}

impl ReadLimits {
    /// Start building a checked read policy from the standard defaults.
    #[must_use]
    pub fn builder() -> ReadLimitsBuilder {
        ReadLimitsBuilder {
            limits: Self::default(),
        }
    }

    /// Maximum accepted compressed package input bytes.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.input_bytes
    }
    /// Maximum non-directory ZIP members.
    #[must_use]
    pub const fn max_archive_members(self) -> usize {
        self.archive_members
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
    pub const fn max_archive_entry_bytes(self) -> u64 {
        self.archive_entry_bytes
    }
    /// Maximum aggregate declared uncompressed ZIP bytes.
    #[must_use]
    pub const fn max_archive_total_bytes(self) -> u64 {
        self.archive_total_bytes
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
    /// Maximum XML events in one relationship part.
    #[must_use]
    pub const fn max_xml_events(self) -> usize {
        self.xml_events
    }
    /// Maximum XML events across relationship parts.
    #[must_use]
    pub const fn max_total_relationship_xml_events(self) -> usize {
        self.total_relationship_xml_events
    }
    /// Maximum relationship XML nesting depth.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.xml_depth
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
            max_files: self.archive_members,
            max_member_name_bytes: self.archive_member_name_bytes,
            max_metadata_bytes: self.archive_metadata_bytes,
            max_compressed_size: self.archive_compressed_bytes,
            max_entry_size: self.archive_entry_bytes,
            max_total_size: self.archive_total_bytes,
        }
    }

    #[allow(
        clippy::unused_self,
        reason = "kept as a method so crate-internal callers write limits.check(...)"
    )]
    pub(crate) fn check(self, resource: ReadResource, actual: u64, maximum: u64) -> Result<()> {
        if actual > maximum {
            return Err(OpcError::ReadLimit {
                resource,
                actual,
                maximum,
            });
        }
        Ok(())
    }

    pub(crate) fn check_input_bytes(self, actual: u64) -> Result<()> {
        self.check(ReadResource::InputBytes, actual, self.input_bytes)
    }
}

impl Default for ReadLimits {
    fn default() -> Self {
        Self {
            input_bytes: 512 * MIB as u64,
            archive_members: 100_000,
            archive_member_name_bytes: 4 * KIB as u64,
            archive_metadata_bytes: 64 * MIB as u64,
            archive_compressed_bytes: 512 * MIB as u64,
            archive_entry_bytes: 512 * MIB as u64,
            archive_total_bytes: 2 * GIB,
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
            xml_events: 1_000_000,
            total_relationship_xml_events: 8_000_000,
            xml_depth: 256,
            xml_attribute_bytes: 64 * KIB,
            relationship_target_bytes: 4 * KIB,
        }
    }
}

/// Builder for [`ReadLimits`].
#[derive(Debug, Clone, Copy)]
pub struct ReadLimitsBuilder {
    limits: ReadLimits,
}

impl ReadLimitsBuilder {
    /// Finalize this checked profile.
    ///
    /// # Errors
    ///
    /// Returns `Err` if the ceilings are inconsistent: relationship parts
    /// exceed archive members, parts exceed archive members, or relationship
    /// target bytes exceed XML attribute bytes.
    pub fn build(self) -> Result<ReadLimits> {
        if self.limits.relationship_parts > self.limits.archive_members {
            return Err(invalid(
                ReadResource::RelationshipParts,
                self.limits.relationship_parts as u64,
            ));
        }
        if self.limits.parts > self.limits.archive_members {
            return Err(invalid(ReadResource::Parts, self.limits.parts as u64));
        }
        if self.limits.relationship_target_bytes > self.limits.xml_attribute_bytes {
            return Err(invalid(
                ReadResource::RelationshipTargetBytes,
                self.limits.relationship_target_bytes as u64,
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
    pub fn max_archive_members(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::ArchiveMembers, value)?;
        self.limits.archive_members = value;
        Ok(self)
    }
    /// Set the ZIP member-name byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_member_name_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::ArchiveMemberNameBytes, value)?;
        self.limits.archive_member_name_bytes = value;
        Ok(self)
    }
    /// Set the ZIP metadata byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_metadata_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::ArchiveMetadataBytes, value)?;
        self.limits.archive_metadata_bytes = value;
        Ok(self)
    }
    /// Set the per-member ZIP compressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_compressed_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::ArchiveCompressedBytes, value)?;
        self.limits.archive_compressed_bytes = value;
        Ok(self)
    }
    /// Set the per-member ZIP uncompressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_entry_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::ArchiveEntryBytes, value)?;
        self.limits.archive_entry_bytes = value;
        Ok(self)
    }
    /// Set the aggregate ZIP uncompressed-byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_archive_total_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::ArchiveTotalBytes, value)?;
        self.limits.archive_total_bytes = value;
        Ok(self)
    }
    /// Set the materialized OPC part count ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_parts(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::Parts, value)?;
        self.limits.parts = value;
        Ok(self)
    }
    /// Set the per-part byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_part_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::PartBytes, value)?;
        self.limits.part_bytes = value;
        Ok(self)
    }
    /// Set the aggregate materialized-part byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_part_bytes(mut self, value: u64) -> Result<Self> {
        validate_bytes(ReadResource::TotalPartBytes, value)?;
        self.limits.total_part_bytes = value;
        Ok(self)
    }
    /// Set the content-types XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_content_types_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::ContentTypesBytes, value)?;
        self.limits.content_types_bytes = value;
        Ok(self)
    }
    /// Set the content-type mapping ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_content_type_mappings(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::ContentTypeMappings, value)?;
        self.limits.content_type_mappings = value;
        Ok(self)
    }
    /// Set the relationship-part ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_parts(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::RelationshipParts, value)?;
        self.limits.relationship_parts = value;
        Ok(self)
    }
    /// Set the per-relationship-part XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_xml_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::RelationshipXmlBytes, value)?;
        self.limits.relationship_xml_bytes = value;
        Ok(self)
    }
    /// Set the aggregate relationship XML byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationship_xml_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::TotalRelationshipXmlBytes, value)?;
        self.limits.total_relationship_xml_bytes = value;
        Ok(self)
    }
    /// Set the relationships-per-part ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationships_per_part(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::RelationshipsPerPart, value)?;
        self.limits.relationships_per_part = value;
        Ok(self)
    }
    /// Set the aggregate relationship ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationships(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::TotalRelationships, value)?;
        self.limits.total_relationships = value;
        Ok(self)
    }
    /// Set the relationship graph-node ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_graph_nodes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::RelationshipGraphNodes, value)?;
        self.limits.relationship_graph_nodes = value;
        Ok(self)
    }
    /// Set the per-part relationship XML event ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_xml_events(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::XmlEvents, value)?;
        self.limits.xml_events = value;
        Ok(self)
    }
    /// Set the aggregate relationship XML event ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_total_relationship_xml_events(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::TotalRelationshipXmlEvents, value)?;
        self.limits.total_relationship_xml_events = value;
        Ok(self)
    }
    /// Set the relationship XML nesting ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_xml_depth(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::XmlDepth, value)?;
        self.limits.xml_depth = value;
        Ok(self)
    }
    /// Set the universal XML attribute byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_xml_attribute_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::XmlAttributeBytes, value)?;
        self.limits.xml_attribute_bytes = value;
        Ok(self)
    }
    /// Set the relationship target byte ceiling.
    ///
    /// # Errors
    ///
    /// Returns `Err` if `value` is zero.
    pub fn max_relationship_target_bytes(mut self, value: usize) -> Result<Self> {
        validate_count(ReadResource::RelationshipTargetBytes, value)?;
        self.limits.relationship_target_bytes = value;
        Ok(self)
    }
}

fn invalid(resource: ReadResource, value: u64) -> OpcError {
    OpcError::InvalidReadLimit { resource, value }
}

fn validate_input(value: u64) -> Result<()> {
    if value == 0 || value > usize::MAX.saturating_sub(1) as u64 {
        return Err(invalid(ReadResource::InputBytes, value));
    }
    Ok(())
}

fn validate_bytes(resource: ReadResource, value: u64) -> Result<()> {
    if value == 0 {
        return Err(invalid(resource, value));
    }
    Ok(())
}

fn validate_count(resource: ReadResource, value: usize) -> Result<()> {
    if value == 0 {
        return Err(invalid(resource, 0));
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
            ReadLimits::builder().max_input_bytes(0),
            Err(OpcError::InvalidReadLimit {
                resource: ReadResource::InputBytes,
                value: 0,
            })
        ));
        assert_eq!(
            ReadLimits::builder()
                .build()
                .unwrap()
                .max_total_part_bytes(),
            512 * MIB as u64
        );
    }
}
