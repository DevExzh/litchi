//! Objects that implement reading and writing OPC packages.
//!
//! This module provides the main `OpcPackage` type, which represents an Open Packaging
//! Convention package in memory. It manages parts, relationships, and provides
//! high-level operations for working with office documents.

use crate::constants::relationship_type;
use crate::content_type::{ContentType, ContentTypeMap};
use crate::error::{OpcError, Result};
use crate::execution::OpenSession;
use crate::leniency::ToleranceReport;
use crate::limits::SourceLimits;
use crate::members::{NonPartMember, PartNameIndex};
use crate::packuri::{CONTENT_TYPES_URI, PACKAGE_URI, PackURI, PartNameConflict};
use crate::part::{Part, PartFactory, PartMetadata};
use crate::payload::PartPayload;
use crate::phys_pkg::{PhysPkgReader, read_limited, read_owned_path_with_limits};
use crate::pkgreader::PackageReader;
use crate::rel::{PreservedRelationshipsXml, Relationships, TargetMode as RelationshipTargetMode};
use std::cmp::Ordering;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Write};
use std::path::Path;
use std::sync::atomic::{AtomicU8, Ordering as AtomicOrdering};
use std::sync::{Arc, OnceLock};

mod content_types;
mod metadata_fragments;
mod relationships;
pub use content_types::{
    ContentTypeEdit, ContentTypeFragment, ContentTypeFragmentEdit, ContentTypeFragmentKey,
    ContentTypeFragmentPlan, ContentTypeFragments, ContentTypesEditPlan,
};
pub use relationships::{
    CanonicalRelationshipsPlan, OwnedRelationships, RelationshipEdit, RelationshipFragment,
    RelationshipFragmentEdit, RelationshipFragmentPlan, RelationshipFragments,
    RelationshipSourcePlan, RelationshipsEditPlan,
};

/// Physical dimensions of an immutable source ZIP retained by an OPC package.
///
/// The metrics are collected from central-directory metadata only. They let a
/// targeted source edit account for opaque non-part members that the writer
/// carries forward without inflating or exposing their payloads. The returned
/// source totals describe the original archive; admission checks below charge
/// only source members the writer will actually copy for the package's current
/// graph, while callers compose newly generated members separately.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RetainedArchiveMetrics {
    members: usize,
    entries: usize,
    aggregate_bytes: u64,
    metadata_bytes: u64,
    opaque_members: usize,
    opaque_entries: usize,
    opaque_bytes: u64,
    known_metadata: Arc<Vec<(String, u64)>>,
}

impl RetainedArchiveMetrics {
    /// Number of non-directory physical members.
    #[must_use]
    pub const fn members(&self) -> usize {
        self.members
    }

    /// Number of central-directory entries, including directory markers.
    #[must_use]
    pub const fn entries(&self) -> usize {
        self.entries
    }

    /// Aggregate declared uncompressed bytes of non-directory members.
    #[must_use]
    pub const fn aggregate_bytes(&self) -> u64 {
        self.aggregate_bytes
    }

    /// Variable central-directory and archive-tail metadata emitted for the
    /// package's current publication image. The fixed central-directory
    /// header is intentionally excluded.
    #[must_use]
    pub const fn metadata_bytes(&self) -> u64 {
        self.metadata_bytes
    }

    /// Number of retained non-directory members outside the OPC projection.
    #[must_use]
    pub const fn opaque_members(&self) -> usize {
        self.opaque_members
    }

    /// Number of retained opaque central-directory entries, including
    /// directory markers.
    #[must_use]
    pub const fn opaque_entries(&self) -> usize {
        self.opaque_entries
    }

    /// Aggregate declared bytes of retained opaque non-directory members.
    #[must_use]
    pub const fn opaque_bytes(&self) -> u64 {
        self.opaque_bytes
    }

    /// Variable metadata bytes in the retained source record for a known OPC
    /// member. The lookup uses the physical member spelling, and returns
    /// `None` for opaque or absent members.
    #[must_use]
    pub fn source_member_metadata(&self, name: &str) -> Option<u64> {
        self.known_metadata
            .iter()
            .find_map(|(member, bytes)| (member == name).then_some(*bytes))
    }
}

/// Count archive-level ZIP metadata which is not attached to one central
/// directory entry.  `metadata_size_hint` covers each entry's variable name,
/// extra-field, and comment fields; the fixed 46-byte central header is not a
/// source-policy resource.  The terminal EOCD comment and (when present) the
/// ZIP64 EOCD extensible data are the remaining variable metadata.
fn archive_tail_metadata_bytes<T: AsRef<[u8]>>(
    archive: &soapberry_zip::ZipSliceArchive<T>,
) -> Result<u64> {
    const ZIP64_EOCD_FIXED_BYTES: u64 = 56;
    const ZIP64_LOCATOR_FIXED_BYTES: u64 = 20;

    let comment_bytes = u64::try_from(archive.comment().as_bytes().len()).map_err(|_| {
        OpcError::ZipError("ZIP EOCD comment length does not fit this platform".to_owned())
    })?;
    if !archive.is_zip64() {
        return Ok(comment_bytes);
    }

    let locator_offset = archive
        .eocd_offset()
        .checked_sub(ZIP64_LOCATOR_FIXED_BYTES)
        .ok_or_else(|| {
            OpcError::ZipError("ZIP64 locator offset underflows the EOCD tail".to_owned())
        })?;
    let fixed_record_end = archive
        .head_eocd_offset()
        .checked_add(ZIP64_EOCD_FIXED_BYTES)
        .ok_or_else(|| OpcError::ZipError("ZIP64 EOCD offset overflows".to_owned()))?;
    let extensible_bytes = locator_offset
        .checked_sub(fixed_record_end)
        .ok_or_else(|| {
            OpcError::ZipError("ZIP64 EOCD extensible data has an invalid span".to_owned())
        })?;
    comment_bytes
        .checked_add(extensible_bytes)
        .ok_or_else(|| OpcError::ZipError("ZIP archive metadata overflows".to_owned()))
}

fn same_optional_manifest_bytes(left: Option<&[u8]>, right: Option<&[u8]>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left == right,
        (Some(_), None) | (None, Some(_)) => false,
    }
}

/// Memoized state for the admitted content-types source binding.
///
/// The binding is invalidated conservatively whenever a mutable package view
/// can change part or relationship metadata. A clone receives its own state,
/// so a staged edit cannot change the source package's memo. `Unknown` keeps
/// the exact source publication behavior: the next query recomputes coverage
/// before deciding whether the retained bytes are still authoritative.
#[derive(Debug)]
struct ContentTypesSourceState(AtomicU8);

impl ContentTypesSourceState {
    const UNKNOWN: u8 = 0;
    const VALID: u8 = 1;
    const INVALID: u8 = 2;

    fn unknown() -> Self {
        Self(AtomicU8::new(Self::UNKNOWN))
    }

    fn load(&self) -> Option<bool> {
        match self.0.load(AtomicOrdering::Relaxed) {
            Self::VALID => Some(true),
            Self::INVALID => Some(false),
            _ => None,
        }
    }

    fn store(&self, valid: bool) {
        self.0.store(
            if valid { Self::VALID } else { Self::INVALID },
            AtomicOrdering::Relaxed,
        );
    }

    fn reset(&self) {
        self.0.store(Self::UNKNOWN, AtomicOrdering::Relaxed);
    }
}

impl Clone for ContentTypesSourceState {
    fn clone(&self) -> Self {
        Self(AtomicU8::new(self.0.load(AtomicOrdering::Relaxed)))
    }
}

/// Options for saving an OPC package.
#[derive(Debug, Clone, Default)]
pub struct SaveOptions {
    /// Typed font-embedding policy; invalid boolean combinations are impossible.
    pub fonts: FontEmbedding,
}

/// Font publication policy used when an Office package is saved.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum FontEmbedding {
    /// Do not discover or publish fonts.
    #[default]
    None,
    /// Publish complete selected font faces.
    Full,
    /// Publish only the glyphs known to be used by the document.
    Subset,
}

#[derive(Debug)]
pub(crate) struct PreservationProvenance {
    pub(crate) members: Vec<SourceMember>,
    pub(crate) parts: HashMap<PackURI, SourcePart>,
    pub(crate) content_types_xml: Arc<Vec<u8>>,
    pub(crate) package_relationships_xml: Arc<PreservedRelationshipsXml>,
}

#[derive(Debug)]
pub(crate) struct SourceMember {
    pub(crate) kind: SourceMemberKind,
}

#[derive(Debug)]
pub(crate) enum SourceMemberKind {
    ContentTypes,
    PackageRelationships,
    Part(PackURI),
    PartRelationships(PackURI),
    Unknown,
}

#[derive(Debug)]
pub(crate) struct SourcePart {
    pub(crate) content_type: String,
    /// The payload the source member carries, captured as the part's own
    /// storage handle. For a deferred part this is the same cell the part
    /// holds, so proving a part untouched costs a pointer comparison and no
    /// decode.
    pub(crate) blob: PartPayload,
    pub(crate) relationships_xml: Arc<PreservedRelationshipsXml>,
    pub(crate) member_present: bool,
    pub(crate) relationships_member_present: bool,
}

/// Apply the same source-payload decision used by targeted publication.
///
/// A deferred current payload is compared through the retained archive's
/// bounded verified reader rather than being inflated into a payload-sized
/// allocation. Equal replacement bytes are safe to copy with the source
/// central record and compressed stream, while a same-length but different
/// replacement is regenerated. Admission and the selected writer therefore
/// make the same exact decision without a conservative length-only branch.
pub(crate) fn source_blob_retained(source_part: &SourcePart, current: Option<&[u8]>) -> bool {
    let Some(current) = current else {
        return true;
    };
    source_part.blob.same_bytes(current).is_ok_and(|same| same)
}

/// A compact semantic fingerprint for one relationship collection.
///
/// The source XML is retained separately.  Keeping the decoded relationship
/// fields here avoids serializing a second XML representation while admitting
/// a source member (attribute escaping can make that representation much
/// larger than the source bytes).
#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipBinding {
    entries: Vec<RelationshipBindingEntry>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipBindingEntry {
    r_id: String,
    reltype: String,
    target_ref: String,
    target_mode: RelationshipTargetMode,
}

impl RelationshipBinding {
    fn from_relationships(relationships: &Relationships) -> Result<Self> {
        let mut entries = Vec::new();
        entries
            .try_reserve_exact(relationships.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship semantic binding entries",
                source,
            })?;
        for relationship in relationships.iter() {
            entries.push(RelationshipBindingEntry {
                r_id: clone_relationship_binding_text(relationship.r_id())?,
                reltype: clone_relationship_binding_text(relationship.reltype())?,
                target_ref: clone_relationship_binding_text(relationship.target_ref())?,
                target_mode: relationship.target_mode(),
            });
        }
        entries.sort_unstable_by(|left, right| left.r_id.cmp(&right.r_id));
        Ok(Self { entries })
    }

    /// Compare a captured semantic binding with the current relationship
    /// collection without cloning any of the current relationship fields.
    ///
    /// `Relationships` is keyed by relationship ID, so this keeps the
    /// comparison linear while retaining the binding's deterministic ID
    /// ordering.  Source capture uses this before admitting a retained XML
    /// allocation under a caller's byte limits.
    fn matches(&self, relationships: &Relationships) -> bool {
        self.entries.len() == relationships.len()
            && self.entries.iter().all(|entry| {
                relationships.get(&entry.r_id).is_some_and(|relationship| {
                    relationship.reltype() == entry.reltype
                        && relationship.target_ref() == entry.target_ref
                        && relationship.target_mode() == entry.target_mode
                })
            })
    }
}

fn check_relationship_capture_limits(
    relationships: &Relationships,
    limits: SourceLimits,
) -> Result<()> {
    limits.check(
        crate::SourceLimitKind::RelationshipsPerPart,
        relationships.len() as u64,
        limits.max_relationships_per_part() as u64,
    )?;
    limits.check(
        crate::SourceLimitKind::TotalRelationships,
        relationships.len() as u64,
        limits.max_total_relationships() as u64,
    )?;
    for relationship in relationships.iter() {
        limits.check(
            crate::SourceLimitKind::XmlAttributeBytes,
            relationship.r_id().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::XmlAttributeBytes,
            relationship.reltype().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::XmlAttributeBytes,
            relationship.target_ref().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::RelationshipTargetBytes,
            relationship.target_ref().len() as u64,
            limits.max_relationship_target_bytes() as u64,
        )?;
    }
    Ok(())
}

fn escaped_relationship_attribute_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |length, character| {
        let encoded = match character {
            '&' => 5,
            '<' | '>' => 4,
            '"' | '\'' => 6,
            _ => character.len_utf8(),
        };
        length.checked_add(encoded).ok_or_else(|| {
            OpcError::InvalidRelationship(
                "canonical relationship XML attribute length overflows".to_owned(),
            )
        })
    })
}

fn check_canonical_relationship_attribute_limits(
    relationships: &Relationships,
    limits: SourceLimits,
) -> Result<()> {
    limits.check(
        crate::SourceLimitKind::XmlAttributeBytes,
        "xmlns".len() as u64 + crate::constants::namespace::OPC_RELATIONSHIPS.len() as u64,
        limits.max_xml_attribute_bytes() as u64,
    )?;
    for relationship in relationships.iter() {
        for (key, value) in [
            ("Id", relationship.r_id()),
            ("Type", relationship.reltype()),
            ("Target", relationship.target_ref()),
        ] {
            let encoded = escaped_relationship_attribute_len(value)?;
            let actual = key.len().checked_add(encoded).ok_or_else(|| {
                OpcError::InvalidRelationship(
                    "canonical relationship XML attribute length overflows".to_owned(),
                )
            })?;
            limits.check(
                crate::SourceLimitKind::XmlAttributeBytes,
                actual as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
        }
        if relationship.target_mode() == RelationshipTargetMode::External {
            let actual = "TargetMode".len() + b"External".len();
            limits.check(
                crate::SourceLimitKind::XmlAttributeBytes,
                actual as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
        }
    }
    Ok(())
}

fn check_canonical_relationship_structure(
    relationships: &Relationships,
    bytes: usize,
    limits: SourceLimits,
) -> Result<()> {
    let events = relationships.len().checked_add(4).ok_or_else(|| {
        OpcError::InvalidRelationship("canonical relationship XML event count overflows".to_owned())
    })?;
    limits.check(
        crate::SourceLimitKind::Events,
        events as u64,
        limits.max_events() as u64,
    )?;
    limits.check(
        crate::SourceLimitKind::TotalRelationshipXmlEvents,
        events as u64,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    let depth = if relationships.is_empty() { 1 } else { 2 };
    limits.check(
        crate::SourceLimitKind::Depth,
        depth,
        limits.max_depth() as u64,
    )?;
    limits.check(
        crate::SourceLimitKind::TotalRelationshipXmlBytes,
        bytes as u64,
        limits.max_total_relationship_xml_bytes() as u64,
    )
}

fn clone_relationship_binding_text(value: &str) -> Result<String> {
    let mut text = String::new();
    text.try_reserve_exact(value.len())
        .map_err(|source| OpcError::Allocation {
            resource: "OPC relationship semantic binding text",
            source,
        })?;
    text.push_str(value);
    Ok(text)
}

fn cmp_ascii_case_insensitive(left: &str, right: &str) -> Ordering {
    for (left, right) in left.bytes().zip(right.bytes()) {
        let ordering = left.to_ascii_lowercase().cmp(&right.to_ascii_lowercase());
        if ordering != Ordering::Equal {
            return ordering;
        }
    }
    left.len().cmp(&right.len())
}

/// Both representations are captured together after relationship admission.
/// Callers cannot authorize arbitrary XML by changing the public edge collection.
#[derive(Debug)]
struct SourceRelationshipsXml {
    bytes: Arc<Vec<u8>>,
    binding: RelationshipBinding,
}

impl PreservedRelationshipsXml {
    /// The provenance of one owner's relationships: the admitted source member
    /// when the package retained one, otherwise the canonical serialization.
    fn from_package(
        package: &OpcPackage,
        owner: &PackURI,
        relationships: &Relationships,
    ) -> Option<Self> {
        if let Some(source) = package.source_relationships_xml.get(owner) {
            Some(Self::Source(Arc::clone(&source.bytes)))
        } else {
            Self::from_relationships(relationships)
        }
    }
}

/// Source-bound or deterministic authored `[Content_Types].xml` metadata.
///
/// The token keeps exact retained XML, or the current deterministic authored
/// XML when no compatible source manifest exists, together with its parsed
/// declarations. A format-level snapshot can carry it across a save/reopen
/// cycle without retaining the whole source archive.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OwnedContentTypes {
    xml: crate::OwnedXmlPart,
    binding: Arc<ContentTypeMap>,
}

impl OwnedContentTypes {
    /// Exact source XML bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.xml.bytes()
    }

    /// Share the exact immutable manifest image without copying its bytes.
    #[must_use]
    pub fn bytes_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml.bytes)
    }

    /// Return the admitted content type resolved for one part name.
    ///
    /// This semantic projection lets a format-level graph patch compare or
    /// compose only the overrides it owns while retaining the source
    /// manifest's unrelated lexical bytes.
    #[must_use]
    pub fn part_content_type(&self, part: &PackURI) -> Option<&str> {
        self.binding.lookup(part).map(ContentType::as_str)
    }

    /// Whether this manifest carries an explicit `<Override>` for a part.
    ///
    /// The resolved type alone cannot distinguish an explicit declaration from
    /// a matching extension default; graph editors use that distinction to
    /// avoid needlessly deleting and appending an unchanged source element.
    #[must_use]
    pub fn has_part_override(&self, part: &PackURI) -> bool {
        self.binding.override_for(part).is_some()
    }

    /// Number of admitted default and override mappings in this manifest.
    #[must_use]
    pub fn mapping_count(&self) -> usize {
        self.binding.mapping_count()
    }

    /// Re-admit this manifest token under a recipient's bounded source policy
    /// without changing a package.
    pub fn preflight_under(&self, limits: SourceLimits) -> Result<()> {
        let limits = self.xml.limits().min(limits);
        self.xml.preflight_under(limits)?;
        let parsed = ContentTypeMap::from_xml(self.bytes(), limits)?;
        if parsed != *self.binding {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types token binding does not match its XML".to_owned(),
            ));
        }
        Ok(())
    }

    /// Remove exact part and relationship-part overrides while preserving all
    /// surrounding manifest bytes. Defaults, comments, ordering, and unrelated
    /// overrides remain source-backed in the returned token.
    pub fn without_parts(&self, parts: &[PackURI], max_output_bytes: usize) -> Result<Self> {
        let xml = content_types::without_part_overrides(&self.xml, parts, max_output_bytes)?;
        if std::ptr::eq(xml.bytes(), self.xml.bytes()) {
            return Ok(self.clone());
        }
        let binding = Arc::new(ContentTypeMap::from_xml(&xml.bytes, xml.limits())?);
        Ok(Self { xml, binding })
    }

    /// Remove only the explicit part overrides for the supplied names.
    ///
    /// The corresponding relationship-member overrides remain in the source
    /// manifest.  This is useful for a graph edit whose owner part remains in
    /// place (for example the host document): its `.rels` member is still
    /// physically present and its authored lexical manifest entry must not be
    /// needlessly removed and recreated.
    pub fn without_part_overrides_only(
        &self,
        parts: &[PackURI],
        max_output_bytes: usize,
    ) -> Result<Self> {
        let xml = content_types::without_part_overrides_only(&self.xml, parts, max_output_bytes)?;
        if std::ptr::eq(xml.bytes(), self.xml.bytes()) {
            return Ok(self.clone());
        }
        let binding = Arc::new(ContentTypeMap::from_xml(&xml.bytes, xml.limits())?);
        Ok(Self { xml, binding })
    }

    /// Append source-preserving explicit overrides for newly added parts.
    /// Existing explicit overrides are rejected so a caller cannot silently
    /// change the declared type of an unrelated part.
    pub fn with_part_overrides(
        &self,
        overrides: &[(&PackURI, &str)],
        max_output_bytes: usize,
    ) -> Result<Self> {
        content_types::preflight_part_overrides(&self.xml, overrides, max_output_bytes)?;
        for (part, _) in overrides {
            if self.binding.override_for(part).is_some() {
                return Err(OpcError::InvalidContentTypesManifest(
                    "content-types override already exists for the part".to_owned(),
                ));
            }
        }
        let xml = content_types::with_part_overrides(&self.xml, overrides, max_output_bytes)?;
        if std::ptr::eq(xml.bytes(), self.xml.bytes()) {
            return Ok(self.clone());
        }
        let binding = Arc::new(ContentTypeMap::from_xml(&xml.bytes, xml.limits())?);
        Ok(Self { xml, binding })
    }
}

/// Main API class for working with OPC packages.
///
/// `OpcPackage` represents an Open Packaging Convention package in memory,
/// providing access to parts, relationships, and package-level operations.
/// Uses efficient data structures and minimal cloning for best performance.
#[allow(
    clippy::module_name_repetitions,
    reason = "OpcPackage is the established public name for the package module's main type."
)]
#[derive(Clone)]
pub struct OpcPackage {
    /// Read policy captured at ingress, retained for format-level edit checks.
    read_limits: SourceLimits,
    /// Package-level relationships
    rels: Relationships,

    /// All parts in the package, indexed by partname
    /// Using Box<dyn Part + Send + Sync> for trait objects to allow different part types
    /// `PackURI` keys avoid string allocations compared to String keys
    parts: HashMap<PackURI, Box<dyn Part + Send + Sync>>,
    /// Retained aggregate length of ordinary part payloads. A `None` value
    /// means a mutable escape hatch may have changed a payload without a
    /// length delta being observable; format owners refresh it at their
    /// publication seam from the existing capture pass.
    part_bytes: Option<u64>,

    /// Exact XML payloads captured from the opened source package. A
    /// deferred part's entry shares that part's payload cell, so the audit
    /// decides an untouched part without decoding it.
    source_xml_parts: HashMap<PackURI, PartPayload>,

    /// Owned source archive retained for exact and targeted publication,
    /// with the digest memo bound to exactly those bytes (change 0751).
    source_archive: Option<retained_archive::RetainedArchive>,

    /// Read limits the owned source archive was admitted under. Meaningful
    /// only while `source_archive` is `Some`; a compressed transfer out of the
    /// archive re-checks the member against the same policy (change 0742).
    source_limits: SourceLimits,

    /// Index of the owned source archive for compressed transfers out of
    /// parts this package materialized eagerly, built on first use and shared
    /// by clones, which share the archive. Deferred parts use the index their
    /// own decode builds (change 0742). `None` without an owned source, so a
    /// package that never transfers pays one small cell at most.
    transfer_index: Option<Arc<TransferIndexCell>>,

    /// Small relationship members survive borrowed ingress without owning the ZIP.
    source_relationships_xml: HashMap<PackURI, Arc<SourceRelationshipsXml>>,

    /// Exact source bytes and parsed declarations for `[Content_Types].xml`.
    /// The bytes are bounded by the structural content-types read limit and
    /// remain available after a reversible part edit.
    source_content_types_xml: Option<Arc<Vec<u8>>>,
    source_content_types: Option<Arc<ContentTypeMap>>,
    content_types_source_state: ContentTypesSourceState,

    /// Clone-local authorization for exact whole-source publication.
    exact_source_authorized: bool,

    /// Whether the package was unmarshaled from an external source. This is
    /// retained even for borrowed ingress, which has no exact-source bytes.
    source_ingress: bool,

    /// Whether this package has had a signature graph established at ingress
    /// or by an explicit authoring operation. New packages may author a
    /// signature graph directly; this marker keeps that case distinct from an
    /// untracked unsigned package while still catching later mutations.
    signature_graph_tracked: bool,

    /// Whether the current signature graph was explicitly handled by a
    /// strip-or-sign operation. Ordinary mutations revoke this policy while a
    /// package remains signed; an explicit unsign remains authorized for later
    /// publication.
    signature_policy_authorized: bool,

    /// Whether a signed source or an API-authored signature graph still needs
    /// an explicit signature edit policy. This remains set after generic
    /// low-level removal empties the graph, so removal cannot masquerade as
    /// signature stripping. Explicit `unsign` clears the requirement.
    signature_policy_required: bool,

    /// Whether the current signature graph was authored through the explicit
    /// signing API. This lets newly authored packages retain the historical
    /// inert-signature construction path while still rejecting later edits to
    /// a graph produced by `sign` or `resign`.
    signature_api_authored: bool,

    /// Source identity used to prove safe targeted publication.
    preservation: Option<Arc<PreservationProvenance>>,

    /// ZIP items the reader found but did not model as parts
    non_part_members: Vec<NonPartMember>,

    /// Reference-compatible repairs applied while the source was read.
    tolerance_report: ToleranceReport,

    /// Save preferences
    save_options: SaveOptions,

    /// The main-document partname, with the relationship it was resolved
    /// from (record 0286). [`OpcPackage::main_document_part`] reuses it only
    /// while the package's main-document relationship still has the same
    /// identifier, target, base and source, so no mutation needs to clear
    /// it; clones share it.
    main_partname: OnceLock<Arc<MainPartname>>,
}

/// A resolved main-document partname and the relationship inputs it was
/// resolved from.
#[derive(Debug)]
struct MainPartname {
    relationship: crate::rel::Relationship,
    partname: PackURI,
}

impl MainPartname {
    /// Whether `relationship` resolves to this partname: the inputs of
    /// [`crate::rel::Relationship::target_partname`] are all equal.
    fn resolves(&self, relationship: &crate::rel::Relationship) -> bool {
        let memo = &self.relationship;
        memo.r_id() == relationship.r_id()
            && memo.target_ref() == relationship.target_ref()
            && memo.base_uri() == relationship.base_uri()
            && memo.source_uri() == relationship.source_uri()
            && memo.target_mode() == relationship.target_mode()
    }
}

impl std::fmt::Debug for OpcPackage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("OpcPackage")
            .field("rels", &self.rels)
            .field("parts_count", &self.parts.len())
            .field("source_xml_parts_count", &self.source_xml_parts.len())
            .field("has_owned_source", &self.source_archive.is_some())
            .field("exact_source_authorized", &self.exact_source_authorized)
            .field("source_ingress", &self.source_ingress)
            .field("signature_graph_tracked", &self.signature_graph_tracked)
            .field(
                "signature_policy_authorized",
                &self.signature_policy_authorized,
            )
            .field("signature_policy_required", &self.signature_policy_required)
            .field("signature_api_authored", &self.signature_api_authored)
            .field("has_preservation_provenance", &self.preservation.is_some())
            .field("non_part_members", &self.non_part_members)
            .field("tolerance_report", &self.tolerance_report)
            .field("save_options", &self.save_options)
            .finish()
    }
}

impl OpcPackage {
    /// Create a new empty OPC package.
    #[must_use]
    pub fn new() -> Self {
        Self {
            read_limits: SourceLimits::default(),
            rels: Relationships::new(PACKAGE_URI.to_string()),
            parts: HashMap::new(),
            part_bytes: Some(0),
            source_xml_parts: HashMap::new(),
            source_relationships_xml: HashMap::new(),
            source_content_types_xml: None,
            source_content_types: None,
            content_types_source_state: ContentTypesSourceState::unknown(),
            source_archive: None,
            source_limits: SourceLimits::default(),
            transfer_index: None,
            exact_source_authorized: false,
            source_ingress: false,
            signature_graph_tracked: false,
            signature_policy_authorized: false,
            signature_policy_required: false,
            signature_api_authored: false,
            preservation: None,
            non_part_members: Vec::new(),
            tolerance_report: ToleranceReport::default(),
            save_options: SaveOptions::default(),
            main_partname: OnceLock::new(),
        }
    }

    /// Return the policy used to read this package.
    ///
    /// Newly authored packages use the default policy. Cloning or editing a
    /// package retains its captured policy so format-level editors can apply
    /// the caller's bounds before staging new payloads. This accessor does
    /// not itself validate edits or change publication behavior.
    #[must_use]
    pub const fn read_limits(&self) -> SourceLimits {
        self.read_limits
    }

    /// ZIP items that were present in the opened archive but are not OPC parts.
    ///
    /// A reader must not reject a package because a ZIP tool left junk in the
    /// archive, but it must not hide the junk either. Each entry names the ZIP
    /// item and why it was not modelled as a part; the bytes stay in the source
    /// archive and are never decompressed.
    #[must_use]
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Re-admit the retained source archive's physical structure under a
    /// tighter operation policy without decoding any member. Per-entry name,
    /// metadata, compressed-size, and uncompressed-size ceilings are checked
    /// only for source members the writer will copy for the current graph;
    /// omitted or regenerated known members are admitted by the caller's
    /// current/prospective metadata checks. Callers use the returned opaque
    /// totals when composing a prospective recipient archive.
    pub fn retained_archive_metrics(
        &self,
        limits: SourceLimits,
    ) -> Result<Option<RetainedArchiveMetrics>> {
        let Some(source) = self.source_archive.as_ref() else {
            return Ok(None);
        };
        let archive = soapberry_zip::ZipArchive::from_slice(source.bytes().as_slice())
            .map_err(OpcError::from)?;
        let declared_entries = usize::try_from(archive.entries_hint()).map_err(|_| {
            OpcError::ZipError("ZIP total entry count does not fit this platform".to_owned())
        })?;
        let mut entries = 0usize;
        let mut members = 0usize;
        let mut metadata_bytes = archive_tail_metadata_bytes(&archive)?;
        let mut aggregate_bytes = 0u64;
        let mut preserved_entries = 0usize;
        let mut preserved_members = 0usize;
        let mut preserved_aggregate_bytes = 0u64;
        let mut opaque_members = 0usize;
        let mut opaque_entries = 0usize;
        let mut opaque_bytes = 0u64;
        // The writer's preservation decision is based on the immutable source
        // classification captured at ingress.  Looking at `self.parts` here
        // is incorrect after an edit: a removed known source part is omitted,
        // while a source member classified as `Unknown` is copied verbatim.
        // Keep the classification indexed by the physical central-directory
        // order so this pass remains metadata-only and follows that exact
        // publication rule.
        let Some(provenance) = self.preservation.as_deref() else {
            return Err(OpcError::PreservationUnavailable {
                reason: "retained source archive has no immutable preservation provenance"
                    .to_owned(),
            });
        };
        let source_members = provenance.members.as_slice();
        let mut known_metadata = Vec::new();
        known_metadata
            .try_reserve_exact(source_members.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC retained known-member metadata",
                source,
            })?;
        let parts_match_source = provenance.parts.len() == self.parts.len()
            && self.iter_parts_undecoded().all(|part| {
                provenance
                    .parts
                    .get(part.partname())
                    .is_some_and(|source| source.content_type == part.content_type())
            });
        let content_types_pristine = parts_match_source
            && self
                .source_content_types_source()?
                .is_some_and(|(source, _)| source == provenance.content_types_xml.as_slice());
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let package_relationships_pristine = self
            .source_relationships_xml(&package_uri, &self.rels)?
            .is_some_and(|source| source == provenance.package_relationships_xml.as_bytes());

        // ZIP central-directory records have a fixed 46-byte header that is
        // not part of the archive metadata budget.  The variable fields are
        // counted by each entry's hint; the terminal EOCD comment and ZIP64
        // extensible data are archive metadata too and must be included once.
        limits.check(
            crate::SourceLimitKind::ArchiveMetadataBytes,
            metadata_bytes,
            limits.max_archive_metadata_bytes(),
        )?;
        for entry_result in archive.entries() {
            let entry = entry_result.map_err(OpcError::from)?;
            entries = entries.checked_add(1).ok_or_else(|| {
                OpcError::ZipError("ZIP central-directory entry count overflows".to_owned())
            })?;
            let source_member = source_members.get(entries - 1).ok_or_else(|| {
                OpcError::ZipError(
                    "ZIP preservation member count is shorter than central directory".to_owned(),
                )
            })?;
            let preserved = match &source_member.kind {
                SourceMemberKind::ContentTypes => content_types_pristine,
                SourceMemberKind::PackageRelationships => package_relationships_pristine,
                SourceMemberKind::Part(partname) => provenance
                    .parts
                    .get(partname)
                    .zip(self.parts.get(partname))
                    .is_some_and(|(source_part, current)| {
                        source_part.member_present
                            && source_blob_retained(source_part, current.decoded_blob())
                    }),
                SourceMemberKind::PartRelationships(partname) => {
                    match provenance.parts.get(partname).zip(self.parts.get(partname)) {
                        Some((source_part, current))
                            if source_part.relationships_member_present =>
                        {
                            self.source_relationships_retained(
                                partname,
                                source_part,
                                current.rels(),
                            )?
                        },
                        _ => false,
                    }
                },
                // Unknown source members are copied by the targeted writer,
                // including when the current OPC graph has a name collision.
                SourceMemberKind::Unknown => true,
            };
            if matches!(&source_member.kind, &SourceMemberKind::Unknown) {
                opaque_entries = opaque_entries.checked_add(1).ok_or_else(|| {
                    OpcError::ZipError("ZIP opaque entry count overflows".to_owned())
                })?;
            }
            if preserved {
                preserved_entries = preserved_entries.checked_add(1).ok_or_else(|| {
                    OpcError::ZipError("ZIP preserved entry count overflows".to_owned())
                })?;
                limits.check(
                    crate::SourceLimitKind::ArchiveMemberNameBytes,
                    entry.file_path().as_ref().len() as u64,
                    limits.max_archive_member_name_bytes(),
                )?;
            }
            let generated = match &source_member.kind {
                SourceMemberKind::ContentTypes | SourceMemberKind::PackageRelationships => {
                    !preserved
                },
                SourceMemberKind::Part(partname) => {
                    !preserved
                        && self.parts.contains_key(partname)
                        && provenance
                            .parts
                            .get(partname)
                            .is_some_and(|source_part| source_part.member_present)
                },
                SourceMemberKind::PartRelationships(partname) => {
                    !preserved
                        && self.parts.get(partname).is_some_and(|part| {
                            !part.rels().is_empty()
                                || self.source_relationships_member_present(partname)
                        })
                },
                SourceMemberKind::Unknown => false,
            };
            let source_entry_metadata = entry.metadata_size_hint();
            if !matches!(&source_member.kind, SourceMemberKind::Unknown)
                && let Ok(name) = std::str::from_utf8(entry.file_path().as_ref())
            {
                known_metadata.push((name.to_owned(), source_entry_metadata));
            }
            let entry_metadata = if generated {
                // RegeneratedEntry has no extra field or comment; its name is
                // the source name retained by the targeted writer.
                entry.file_path().as_ref().len() as u64
            } else {
                source_entry_metadata
            };
            if preserved || generated {
                metadata_bytes = metadata_bytes.checked_add(entry_metadata).ok_or_else(|| {
                    OpcError::ZipError("ZIP central-directory metadata overflows".to_owned())
                })?;
                limits.check(
                    crate::SourceLimitKind::ArchiveMetadataBytes,
                    metadata_bytes,
                    limits.max_archive_metadata_bytes(),
                )?;
            }
            if entry.is_dir() {
                continue;
            }
            members = members
                .checked_add(1)
                .ok_or_else(|| OpcError::ZipError("ZIP member count overflows".to_owned()))?;
            if preserved {
                preserved_members = preserved_members.checked_add(1).ok_or_else(|| {
                    OpcError::ZipError("ZIP preserved member count overflows".to_owned())
                })?;
                limits.check(
                    crate::SourceLimitKind::Members,
                    preserved_members as u64,
                    limits.max_members() as u64,
                )?;
                limits.check(
                    crate::SourceLimitKind::ArchiveCompressedBytes,
                    entry.compressed_size_hint(),
                    limits.max_archive_compressed_bytes(),
                )?;
                limits.check(
                    crate::SourceLimitKind::EntryBytes,
                    entry.uncompressed_size_hint(),
                    limits.max_entry_bytes(),
                )?;
            }
            aggregate_bytes = aggregate_bytes
                .checked_add(entry.uncompressed_size_hint())
                .ok_or_else(|| {
                    OpcError::ZipError("ZIP aggregate uncompressed bytes overflow".to_owned())
                })?;
            if preserved {
                preserved_aggregate_bytes = preserved_aggregate_bytes
                    .checked_add(entry.uncompressed_size_hint())
                    .ok_or_else(|| {
                        OpcError::ZipError("ZIP preserved aggregate bytes overflow".to_owned())
                    })?;
                limits.check(
                    crate::SourceLimitKind::AggregateBytes,
                    preserved_aggregate_bytes,
                    limits.max_aggregate_bytes(),
                )?;
            }
            if matches!(&source_member.kind, &SourceMemberKind::Unknown) {
                opaque_members = opaque_members.checked_add(1).ok_or_else(|| {
                    OpcError::ZipError("ZIP opaque member count overflows".to_owned())
                })?;
                opaque_bytes = opaque_bytes
                    .checked_add(entry.uncompressed_size_hint())
                    .ok_or_else(|| {
                        OpcError::ZipError("ZIP opaque member bytes overflow".to_owned())
                    })?;
            }
        }
        if entries != declared_entries {
            return Err(OpcError::ZipError(
                "ZIP central-directory entry count differs from EOCD".to_owned(),
            ));
        }
        if source_members.len() != entries {
            return Err(OpcError::ZipError(
                "ZIP preservation member count differs from central directory".to_owned(),
            ));
        }
        // The ordinary writer always emits the package-level relationships
        // member, even when the source archive omitted `_rels/.rels` and the
        // targeted preservation path therefore falls back to ordinary
        // publication.  Charge the generated member name exactly once.  A
        // source member that was present is already in `known_metadata`, so
        // this also keeps the projection aligned with the writer's central
        // directory without double-counting it.
        let package_relationship_name = package_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;
        if !known_metadata
            .iter()
            .any(|(name, _)| name == package_relationship_name.membername())
        {
            metadata_bytes = metadata_bytes
                .checked_add(package_relationship_name.membername().len() as u64)
                .ok_or_else(|| {
                    OpcError::ZipError("ZIP central-directory metadata overflows".to_owned())
                })?;
            limits.check(
                crate::SourceLimitKind::ArchiveMetadataBytes,
                metadata_bytes,
                limits.max_archive_metadata_bytes(),
            )?;
        }
        limits.check(
            crate::SourceLimitKind::ArchiveTotalEntries,
            preserved_entries as u64,
            limits.max_archive_total_entries() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::Members,
            preserved_members as u64,
            limits.max_members() as u64,
        )?;
        // A targeted publication appends newly authored parts and relationship
        // members after the retained source namespace. Their generated
        // central records have only the member name as variable metadata, so
        // account for them here as part of the same recipient summary.
        for part in self.iter_parts_undecoded() {
            if !known_metadata
                .iter()
                .any(|(name, _)| name == part.partname().membername())
            {
                metadata_bytes = metadata_bytes
                    .checked_add(part.partname().membername().len() as u64)
                    .ok_or_else(|| {
                        OpcError::ZipError("ZIP central-directory metadata overflows".to_owned())
                    })?;
                limits.check(
                    crate::SourceLimitKind::ArchiveMetadataBytes,
                    metadata_bytes,
                    limits.max_archive_metadata_bytes(),
                )?;
            }
            if !part.rels().is_empty() {
                let relationship_name = part
                    .partname()
                    .rels_uri()
                    .map_err(OpcError::InvalidPackUri)?;
                if !known_metadata
                    .iter()
                    .any(|(name, _)| name == relationship_name.membername())
                {
                    metadata_bytes = metadata_bytes
                        .checked_add(relationship_name.membername().len() as u64)
                        .ok_or_else(|| {
                            OpcError::ZipError(
                                "ZIP central-directory metadata overflows".to_owned(),
                            )
                        })?;
                    limits.check(
                        crate::SourceLimitKind::ArchiveMetadataBytes,
                        metadata_bytes,
                        limits.max_archive_metadata_bytes(),
                    )?;
                }
            }
        }
        Ok(Some(RetainedArchiveMetrics {
            members,
            entries,
            aggregate_bytes,
            metadata_bytes,
            opaque_members,
            opaque_entries,
            opaque_bytes,
            known_metadata: Arc::new(known_metadata),
        }))
    }

    /// Replace the reader-reported ZIP members that are not OPC parts.
    ///
    /// This is crate-private because only package ingress paths can establish
    /// the classification; mutable package callers cannot manufacture it.
    pub(crate) fn set_non_part_members(&mut self, members: Vec<NonPartMember>) {
        self.non_part_members = members;
    }

    /// Reference-compatible repairs applied while this package was read.
    ///
    /// The report describes the opened source, in discovery order: an empty
    /// report means the package conformed. Repairs are applied under the
    /// default contract of [`crate::leniency`] — ZIP framing, XML
    /// well-formedness, and read limits stay hard errors. A newly created
    /// package reports clean, and edits never add to the report. Format crates
    /// fold this report into their own document-level report.
    #[must_use]
    pub fn tolerance_report(&self) -> &ToleranceReport {
        &self.tolerance_report
    }

    /// Replace the reader-reported tolerance report.
    ///
    /// Crate-private for the same reason as [`Self::set_non_part_members`]:
    /// only package ingress can establish what was repaired.
    pub(crate) fn set_tolerance_report(&mut self, report: ToleranceReport) {
        self.tolerance_report = report;
    }

    pub(crate) fn source_relationships_member_present(&self, partname: &PackURI) -> bool {
        self.source_relationships_xml.contains_key(partname)
    }

    /// Whether an admitted source archive carried a physical relationships
    /// member for `owner`, even when that member's relationship collection is
    /// empty. This metadata-only projection is used by bounded publication
    /// planning to retain the writer's empty-member decision without decoding
    /// any payload.
    #[doc(hidden)]
    #[must_use]
    pub fn has_retained_relationship_member(&self, owner: &PackURI) -> bool {
        self.source_relationships_member_present(owner)
    }

    fn check_source_content_types_limits(
        &self,
        name: &PackURI,
        bytes: &[u8],
        source: &ContentTypeMap,
        limits: SourceLimits,
    ) -> Result<()> {
        limits.check(
            crate::SourceLimitKind::ContentTypesBytes,
            bytes.len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::PartBytes,
            bytes.len() as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            crate::SourceLimitKind::Parts,
            self.parts.len() as u64,
            limits.max_parts() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::ContentTypeMappings,
            source.mapping_count() as u64,
            limits.max_content_type_mappings() as u64,
        )?;
        crate::OwnedXmlPart::check_capture_size(name, bytes.len(), limits)
    }

    pub(crate) fn source_content_types_source(&self) -> Result<Option<(&[u8], &ContentTypeMap)>> {
        let (bytes, map) = match (
            self.source_content_types_xml.as_ref(),
            self.source_content_types.as_ref(),
        ) {
            (Some(bytes), Some(map)) => (bytes.as_slice(), map.as_ref()),
            _ => return Ok(None),
        };
        if let Some(valid) = self.content_types_source_state.load() {
            return Ok(valid.then_some((bytes, map)));
        }
        let matches = self.source_content_types_matches_current_parts(map)?;
        self.content_types_source_state.store(matches);
        Ok(matches.then_some((bytes, map)))
    }

    /// Whether `content_type` is exactly what the retained source manifest
    /// declared for `partname`.
    ///
    /// Source content types are admitted verbatim at ingress even when they
    /// break the MIME grammar (LENIENCY_AUDIT 8.12, recorded in the tolerance
    /// report), so a regenerated manifest may carry such a value forward
    /// unchanged; content types the caller authors keep the grammar check.
    pub(crate) fn content_type_is_from_source(
        &self,
        partname: &PackURI,
        content_type: &str,
    ) -> bool {
        self.source_content_types.as_ref().is_some_and(|source| {
            source
                .lookup(partname)
                .is_some_and(|declared| declared.as_str() == content_type)
        })
    }

    /// Whether the retained source manifest is still `allocation`, the one
    /// the open captured. A token installed since then
    /// ([`Self::try_replace_content_types`] or a batch add) replaces the
    /// allocation, so this is `false` even when that token still describes
    /// every part exactly as the source manifest did.
    pub(crate) fn retains_source_content_types(&self, allocation: &Arc<Vec<u8>>) -> bool {
        self.source_content_types_xml
            .as_ref()
            .is_some_and(|current| Arc::ptr_eq(current, allocation))
    }

    fn source_content_types_matches_current_parts(&self, source: &ContentTypeMap) -> Result<bool> {
        if self.parts.values().any(|part| {
            source
                .lookup(part.partname())
                .is_none_or(|content_type| content_type.as_str() != part.content_type())
        }) {
            return Ok(false);
        }

        // Index borrowed physical part names once. Relationship overrides are
        // uncommon; derive only their owner names, rather than allocating a
        // relationship URI for every part on every manifest read.
        let mut current_names = Vec::new();
        current_names
            .try_reserve_exact(self.parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC content-types current part-name index",
                source,
            })?;
        current_names.extend(self.parts.keys());
        current_names.sort_unstable_by(|left, right| {
            cmp_ascii_case_insensitive(left.as_str(), right.as_str())
        });

        for (partname, content_type) in source.overrides() {
            if current_names
                .binary_search_by(|current| {
                    cmp_ascii_case_insensitive(current.as_str(), partname.as_str())
                })
                .is_ok()
            {
                continue;
            }
            if !content_type
                .as_str()
                .eq_ignore_ascii_case(crate::constants::content_type::OPC_RELATIONSHIPS)
            {
                return Ok(false);
            }
            if partname.as_str().eq_ignore_ascii_case("/_rels/.rels") {
                let root = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
                if self.rels.is_empty() && !self.source_relationships_member_present(&root) {
                    return Ok(false);
                }
                continue;
            }
            let Some((directory, filename)) = partname.as_str().rsplit_once('/') else {
                return Ok(false);
            };
            let Some((parent, marker)) = directory.rsplit_once('/') else {
                return Ok(false);
            };
            if !marker.eq_ignore_ascii_case("_rels")
                || !filename
                    .get(filename.len().saturating_sub(5)..)
                    .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".rels"))
                || filename.len() <= 5
            {
                return Ok(false);
            }
            // The ASCII suffix proves this truncation is a UTF-8 boundary.
            let filename = &filename[..filename.len() - 5];
            let mut owner_name = String::new();
            owner_name
                .try_reserve_exact(parent.len() + 1 + filename.len())
                .map_err(|source| OpcError::Allocation {
                    resource: "OPC content-types relationship override owner",
                    source,
                })?;
            owner_name.push_str(parent);
            owner_name.push('/');
            owner_name.push_str(filename);
            let Ok(index) = current_names.binary_search_by(|current| {
                cmp_ascii_case_insensitive(current.as_str(), &owner_name)
            }) else {
                return Ok(false);
            };
            let owner = current_names[index];
            let part = &self.parts[owner];
            if part.rels().is_empty() && !self.source_relationships_member_present(owner) {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// Record the result that is already known when a new part's declared
    /// type is absent from the retained source binding.  The full coverage
    /// walk remains necessary for a type that is covered: relationship
    /// overrides and other existing graph changes can still affect source
    /// publication.  This fast path only records a definitive invalidation.
    fn mark_content_types_invalid_if_uncovered(&self, part: &dyn Part) {
        if self.source_content_types.as_ref().is_some_and(|source| {
            source
                .lookup(part.partname())
                .is_none_or(|content_type| content_type.as_str() != part.content_type())
        }) {
            self.content_types_source_state.store(false);
        }
    }

    /// Capture the current `[Content_Types].xml` publication view.
    ///
    /// Compatible admitted source bytes are retained exactly. After an edit
    /// that invalidates source coverage, the view is deterministic authored
    /// XML for the current part/type set. Newly authored packages use the
    /// same deterministic view. Callers cannot construct the token from
    /// arbitrary XML, and replacement is accepted only for the current
    /// expected view.
    /// Share the retained content-types bytes when they still describe this
    /// graph. This is a read-only projection of admitted source bytes, not a
    /// source token or a new admission under a caller's parsing policy.
    ///
    /// Structural edits can invalidate coverage, in which case this returns
    /// `None`; [`Self::authored_content_types_xml`] provides the bounded
    /// deterministic fallback. Changing the returned allocation with
    /// `Arc::make_mut` cannot change this package or authorize replacement.
    pub fn retained_content_types_xml(&self) -> Result<Option<Arc<Vec<u8>>>> {
        if self.source_content_types_source()?.is_some() {
            Ok(self.source_content_types_xml.clone())
        } else {
            Ok(None)
        }
    }

    /// Share the admitted source content-types bytes even when a structural
    /// edit has made their binding incompatible with the current part graph.
    /// This is provenance only; publication uses the authored fallback until
    /// the graph matches the source binding again, and callers still need the
    /// validating replacement API to install new bytes.
    #[must_use]
    pub fn content_types_source_bytes(&self) -> Option<Arc<Vec<u8>>> {
        self.source_content_types_xml.clone()
    }

    /// Serialize the typed content-type graph under required output limits.
    /// This deterministic authored view does not retain producer XML and does
    /// not grant source-replacement authority. Use the retained projection
    /// first when an exact publication view is needed.
    pub fn authored_content_types_xml(&self, limits: SourceLimits) -> Result<Arc<Vec<u8>>> {
        let limits = self.read_limits.min(limits);
        Ok(Arc::new(
            crate::pkgwriter::authored_content_types_xml_with_limits(self, limits)?,
        ))
    }

    /// Capture the current `[Content_Types].xml` publication view under an
    /// explicit bounded read policy. Retained source bytes and declarations
    /// are checked before compatibility indexing; authored fallback uses the
    /// same limits for its two-pass size and mapping preflight.
    pub fn source_content_types(&self, limits: SourceLimits) -> Result<OwnedContentTypes> {
        let limits = self.read_limits.min(limits);
        let name = PackURI::new(CONTENT_TYPES_URI).map_err(OpcError::InvalidPackUri)?;
        crate::OwnedXmlPart::check_capture_member_name(&name, limits)?;
        let retained = match (
            self.source_content_types_xml.as_ref(),
            self.source_content_types.as_ref(),
        ) {
            (Some(bytes), Some(binding)) => {
                self.check_source_content_types_limits(&name, bytes, binding, limits)?;
                if self.source_content_types_matches_current_parts(binding)? {
                    Some((Arc::clone(bytes), Arc::clone(binding)))
                } else {
                    None
                }
            },
            _ => None,
        };
        let (bytes, binding) = match retained {
            Some(source) => source,
            None => {
                let bytes = Arc::new(crate::pkgwriter::authored_content_types_xml_with_limits(
                    self, limits,
                )?);
                limits.check(
                    crate::SourceLimitKind::ContentTypesBytes,
                    bytes.len() as u64,
                    limits.max_content_types_bytes() as u64,
                )?;
                limits.check(
                    crate::SourceLimitKind::PartBytes,
                    bytes.len() as u64,
                    limits.max_part_bytes(),
                )?;
                let binding = Arc::new(ContentTypeMap::from_xml(bytes.as_slice(), limits)?);
                (bytes, binding)
            },
        };
        let xml = crate::OwnedXmlPart::capture_with_limits(
            name,
            crate::constants::content_type::XML.to_owned(),
            bytes,
            limits,
        )?;
        Ok(OwnedContentTypes { xml, binding })
    }

    /// Replace manifest XML decoded from a durable format-level patch.
    ///
    /// `expected` must match the current publication view. Replacement bytes
    /// are untrusted: bounded XML and content-type parsing and the publication
    /// audit run before mutation. This does not create a source token from raw
    /// bytes or bypass signature policy. The format owner remains responsible
    /// for applying the corresponding part/type dependency closure and reserving
    /// package-wide aggregate budgets. This is exact-image restoration; use
    /// structural metadata plans when transporting edits onto a divergent source.
    pub fn try_replace_content_types_bytes(
        &mut self,
        expected: &[u8],
        replacement: Arc<Vec<u8>>,
        limits: SourceLimits,
    ) -> Result<bool> {
        let limits = self.read_limits.min(limits);
        let current = self.source_content_types(limits)?;
        if current.bytes() != expected {
            return Err(OpcError::InvalidContentTypesManifest(
                "stale content-types source replacement".to_owned(),
            ));
        }
        let source_differs = self
            .source_content_types_xml
            .as_ref()
            .is_some_and(|source| source.as_slice() != replacement.as_slice());
        let signed = self.is_signed() || self.requires_signature_edit_policy();
        if current.bytes() == replacement.as_slice() && (!source_differs || signed) {
            return Ok(false);
        }
        if signed {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let name = PackURI::new(CONTENT_TYPES_URI).map_err(OpcError::InvalidPackUri)?;
        crate::OwnedXmlPart::check_capture_size(&name, replacement.len(), limits)?;
        let binding = Arc::new(ContentTypeMap::from_xml(&replacement, limits)?);
        let xml = crate::OwnedXmlPart::capture_with_limits(
            name,
            crate::constants::content_type::XML.to_owned(),
            replacement,
            limits,
        )?;
        self.try_replace_content_types(expected, &OwnedContentTypes { xml, binding }, limits)
    }

    /// Replace an exact source-manifest snapshot.
    ///
    /// The caller owns the format-level dependency closure. Publication only
    /// reuses the token once the current part/type coverage agrees with its
    /// parsed declarations, so restoring the token before re-adding a removed
    /// part remains safe and deterministic.
    /// Replace an exact source-manifest snapshot under an explicit bounded
    /// read policy.
    ///
    /// The current manifest is checked before replacement validation or any
    /// package mutation. A changed signed package is refused before parsing
    /// or retaining replacement metadata, and the replacement's XML is
    /// admitted under the same limits used for the current source check.
    pub fn try_replace_content_types(
        &mut self,
        expected: &[u8],
        replacement: &OwnedContentTypes,
        limits: SourceLimits,
    ) -> Result<bool> {
        let limits = self.read_limits.min(limits);
        let current = self.source_content_types(limits)?;
        if current.bytes() != expected {
            return Err(OpcError::InvalidContentTypesManifest(
                "stale content-types source replacement".to_owned(),
            ));
        }
        if replacement.xml.name.as_str() != CONTENT_TYPES_URI
            || replacement.xml.content_type != crate::constants::content_type::XML
        {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types replacement has an invalid owner or content type".to_owned(),
            ));
        }
        let source_differs = self
            .source_content_types_xml
            .as_ref()
            .is_some_and(|source| source.as_slice() != replacement.bytes());
        let signed = self.is_signed() || self.requires_signature_edit_policy();
        if current.bytes() == replacement.bytes() && (!source_differs || signed) {
            return Ok(false);
        }
        if signed {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        limits.check(
            crate::SourceLimitKind::ContentTypesBytes,
            replacement.bytes().len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            crate::SourceLimitKind::ContentTypeMappings,
            replacement.binding.mapping_count() as u64,
            limits.max_content_type_mappings() as u64,
        )?;
        crate::OwnedXmlPart::check_capture_size(
            &replacement.xml.name,
            replacement.bytes().len(),
            limits,
        )?;
        // The token's binding is the exact semantic companion of its bytes.
        // Re-parse under the caller's policy to enforce XML event, depth, and
        // attribute ceilings without replacing that retained binding.
        let parsed = ContentTypeMap::from_xml(replacement.bytes(), limits)?;
        if parsed != *replacement.binding {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types replacement binding does not match its XML".to_owned(),
            ));
        }
        // The token becomes the retained manifest, which the writer publishes
        // without its audit (change 0665). Its bytes came from the caller, so
        // audit them before any mutation (ADR 0006).
        verified_metadata_token(
            &replacement.xml.name,
            Arc::clone(&replacement.xml.bytes),
            limits.max_content_types_bytes(),
        )?;
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.source_content_types_xml = Some(Arc::clone(&replacement.xml.bytes));
        self.source_content_types = Some(Arc::clone(&replacement.binding));
        Ok(true)
    }

    pub(crate) fn source_relationships_xml(
        &self,
        owner: &PackURI,
        relationships: &Relationships,
    ) -> Result<Option<&[u8]>> {
        Ok(self
            .source_relationships_xml
            .get(owner)
            .filter(|source| source.binding.matches(relationships))
            .map(|source| source.bytes.as_slice()))
    }

    /// Whether a current relationship collection will be copied with the
    /// source central record by targeted publication.
    ///
    /// The source capture settles the untouched case without serialization.
    /// A mutate-then-revert collection keeps its semantic source binding and
    /// is likewise copied byte-for-byte. Only a collection with no matching
    /// source binding needs one canonical serialization for the final byte
    /// comparison the writer performs.
    pub(crate) fn source_relationships_retained(
        &self,
        owner: &PackURI,
        source_part: &SourcePart,
        relationships: &Relationships,
    ) -> Result<bool> {
        if !source_part.relationships_member_present {
            return Ok(false);
        }
        if relationships
            .source_capture()
            .is_some_and(|capture| Arc::ptr_eq(capture, &source_part.relationships_xml))
        {
            return Ok(true);
        }
        if let Some(source) = self.source_relationships_xml(owner, relationships)? {
            return Ok(source_part.relationships_xml.as_bytes() == source);
        }
        Ok(source_part.relationships_xml.as_bytes() == relationships.try_to_xml_bytes()?)
    }

    /// Whether `part` will use the retained source member and central record
    /// during targeted publication.
    #[doc(hidden)]
    pub fn source_part_member_retained(&self, part: &dyn Part) -> bool {
        self.preservation
            .as_deref()
            .and_then(|provenance| provenance.parts.get(part.partname()))
            .is_some_and(|source_part| {
                source_part.member_present && source_blob_retained(source_part, part.decoded_blob())
            })
    }

    /// Whether the current relationship collection will use its retained
    /// source member and central record during targeted publication.
    #[doc(hidden)]
    pub fn source_relationship_member_retained(
        &self,
        owner: &PackURI,
        relationships: &Relationships,
    ) -> Result<bool> {
        let Some(source_part) = self
            .preservation
            .as_deref()
            .and_then(|provenance| provenance.parts.get(owner))
        else {
            return Ok(false);
        };
        self.source_relationships_retained(owner, source_part, relationships)
    }

    /// Whether the exact source content-types bytes can be copied for a
    /// projected publication image.
    #[doc(hidden)]
    pub fn source_content_types_member_retained(&self, bytes: &[u8]) -> bool {
        self.preservation
            .as_deref()
            .is_some_and(|provenance| provenance.content_types_xml.as_slice() == bytes)
    }

    /// Compare retained content-type spelling without serializing fallback XML.
    /// This is a provenance comparison; callers must separately compare the
    /// typed graph when both sides would use authored fallback metadata.
    pub fn same_retained_content_types(&self, other: &Self) -> Result<bool> {
        // A clone with the same manifest allocation and the same coverage
        // shape has the same retained/fallback decision. Avoid rebuilding and
        // sorting two part-name indexes for a payload-only edit. Relationship
        // presence matters because a producer may override a .rels member.
        if let (Some(left), Some(right), Some(left_map), Some(right_map)) = (
            &self.source_content_types_xml,
            &other.source_content_types_xml,
            &self.source_content_types,
            &other.source_content_types,
        ) && Arc::ptr_eq(left, right)
            && Arc::ptr_eq(left_map, right_map)
            && self.parts.len() == other.parts.len()
            && self.rels.is_empty() == other.rels.is_empty()
            && self.parts.iter().all(|(name, part)| {
                other.parts.get(name).is_some_and(|candidate| {
                    part.partname() == candidate.partname()
                        && part.content_type() == candidate.content_type()
                        && part.rels().is_empty() == candidate.rels().is_empty()
                        && self.source_relationships_member_present(name)
                            == other.source_relationships_member_present(name)
                })
            })
        {
            let root = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
            if self.source_relationships_member_present(&root)
                == other.source_relationships_member_present(&root)
            {
                return Ok(true);
            }
        }
        Ok(same_optional_manifest_bytes(
            self.source_content_types_source()?.map(|(bytes, _)| bytes),
            other.source_content_types_source()?.map(|(bytes, _)| bytes),
        ))
    }

    /// Return the retained relationship-member bytes when their semantic
    /// binding still matches the current typed graph.
    ///
    /// This is the zero-copy provenance path for format patch capture. A
    /// caller that needs a canonical member for an authored or changed graph
    /// should use [`Self::source_relationships`]. The boolean reports whether
    /// the physical relationship member is present; a source member with no
    /// edges is therefore distinguishable from an absent part member.
    pub fn retained_relationships_xml(
        &self,
        owner: &PackURI,
    ) -> Result<Option<(Arc<Vec<u8>>, bool)>> {
        let (owner_ref, relationships) = if owner.as_str() == PACKAGE_URI {
            (owner, self.rels())
        } else {
            let Some(part) = self.part_metadata(owner) else {
                return Ok(None);
            };
            (part.partname(), part.rels())
        };
        let Some(source) = self
            .source_relationships_xml
            .get(owner_ref)
            .filter(|source| source.binding.matches(relationships))
        else {
            return Ok(None);
        };
        let member_present = owner_ref.as_str() == PACKAGE_URI
            || !relationships.is_empty()
            || self.source_relationships_member_present(owner_ref);
        Ok(Some((Arc::clone(&source.bytes), member_present)))
    }

    /// Compare one owner's retained relationship spelling and physical presence.
    /// Typed relationship equality is a separate caller check; two absent
    /// retained tokens may have different authored fallback graphs.
    pub fn same_retained_relationships(&self, other: &Self, owner: &PackURI) -> Result<bool> {
        let left = if owner.as_str() == PACKAGE_URI {
            Some(&self.rels)
        } else {
            self.part_metadata(owner).map(|part| part.rels())
        };
        let right = if owner.as_str() == PACKAGE_URI {
            Some(&other.rels)
        } else {
            other.part_metadata(owner).map(|part| part.rels())
        };
        match (left, right) {
            (None, None) => Ok(true),
            (Some(left), Some(right)) => Ok((!left.is_empty()
                || self.source_relationships_member_present(owner))
                == (!right.is_empty() || other.source_relationships_member_present(owner))
                && (!self.source_relationships_member_present(owner)
                    && !other.source_relationships_member_present(owner)
                    || left.same_source_capture(right)
                    || same_optional_manifest_bytes(
                        self.source_relationships_xml(owner, left)?,
                        other.source_relationships_xml(owner, right)?,
                    ))),
            _ => Ok(false),
        }
    }

    /// Whether the retained relationship-member captures have the same
    /// publication bytes. Clones share each capture allocation; comparing the
    /// map once lets patch capture skip an owner-by-owner provenance walk when
    /// no raw relationship member was replaced.
    #[doc(hidden)]
    #[must_use]
    pub fn same_relationship_source_captures(&self, other: &Self) -> bool {
        self.source_relationships_xml.len() == other.source_relationships_xml.len()
            && self.source_relationships_xml.iter().all(|(owner, left)| {
                other
                    .source_relationships_xml
                    .get(owner)
                    .is_some_and(|right| {
                        Arc::ptr_eq(&left.bytes, &right.bytes)
                            || left.bytes.as_slice() == right.bytes.as_slice()
                    })
            })
    }

    /// Compare the retained publication manifests without serializing them.
    ///
    /// The typed OPC graph does not include `[Content_Types].xml` or the
    /// relationship members as ordinary parts.  When either package still
    /// has a valid retained source token, compare that member's exact bytes;
    /// when both sides have no token they will use the same deterministic
    /// authored fallback after the caller has established graph equality.
    /// A retained token on only one side, or different retained bytes, is a
    /// conservative inequality because the source may carry producer markup
    /// outside the typed graph.
    pub fn same_retained_publication_manifests(&self, other: &Self) -> Result<bool> {
        if !same_optional_manifest_bytes(
            self.source_content_types_source()?.map(|(bytes, _)| bytes),
            other.source_content_types_source()?.map(|(bytes, _)| bytes),
        ) {
            return Ok(false);
        }

        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        if !same_optional_manifest_bytes(
            self.source_relationships_xml(&package_uri, &self.rels)?,
            other.source_relationships_xml(&package_uri, &other.rels)?,
        ) {
            return Ok(false);
        }

        for metadata in self.iter_parts() {
            let Some(other_metadata) = other.part_metadata(metadata.partname()) else {
                return Ok(false);
            };
            if !same_optional_manifest_bytes(
                self.source_relationships_xml(metadata.partname(), metadata.rels())?,
                other.source_relationships_xml(other_metadata.partname(), other_metadata.rels())?,
            ) {
                return Ok(false);
            }
        }
        Ok(true)
    }

    /// Set save options for the package.
    pub fn set_save_options(&mut self, options: SaveOptions) {
        self.revoke_exact_source();
        self.save_options = options;
    }

    /// Get current save options.
    #[must_use]
    pub fn save_options(&self) -> &SaveOptions {
        &self.save_options
    }

    /// Whether this package still represents its unmodified owned source.
    ///
    /// This requires both owned ingress (built-in parts and default save
    /// preferences) and unrevoked exact-source authorization. Raw mutable
    /// access, legacy mutators and save-option changes revoke it for that
    /// package clone, including failures and no-ops. Transactional token APIs
    /// may preserve authorization when source checks refuse or detect an
    /// unchanged value before mutation. Borrowed ingress and newly authored
    /// packages return `false`. This query neither exposes source bytes nor
    /// restores authorization.
    #[must_use]
    pub fn is_unmodified_owned_source(&self) -> bool {
        self.exact_source().is_some()
    }

    /// Whether every part of this package is one of `litchi-opc`'s own part
    /// types, [`BlobPart`](crate::BlobPart) or [`XmlPart`](crate::XmlPart).
    ///
    /// Such a package behaves like the reopen of its serialization. A
    /// caller-defined [`Part`] implementation may not: it can refuse a
    /// content-type change, observe writes or count relationship references
    /// its own way, and reopening its bytes yields a built-in part in its
    /// place. A package holding one answers `false`; an unmodified owned
    /// source holds only built-in parts (change 0742).
    #[must_use]
    pub fn holds_only_built_in_parts(&self) -> bool {
        self.parts
            .values()
            .all(|part| part.is_built_in(crate::part::Seal(())))
    }

    /// SHA-256 of the exact owned source archive, while the exact-source
    /// authorization is intact.
    ///
    /// For such a package [`Self::write_to`] writes exactly the retained
    /// archive and nothing else, so this is the digest of the bytes the
    /// package publishes. It is computed on first request, at most once per
    /// owned ingress, and shared by every clone of that ingress, because the
    /// clones share the archive: the memo is created together with the
    /// archive by one private constructor, is filled only from those bytes,
    /// and the archive is never mutated (change 0751). A second caller
    /// therefore reads the digest instead of hashing the archive again.
    ///
    /// Returns `None` exactly when [`Self::exact_source_shared`] does: for a
    /// package that was not opened from owned bytes, and for one whose
    /// exact-source authorization has been revoked by an edit. Like that
    /// handle, a digest already taken is not evidence about the package's
    /// state after a later edit.
    #[must_use]
    pub fn exact_source_sha256(&self) -> Option<[u8; 32]> {
        if !self.exact_source_authorized {
            return None;
        }
        self.source_archive
            .as_ref()
            .map(retained_archive::RetainedArchive::sha256)
    }

    /// Length of the exact owned source archive, while the exact-source
    /// authorization is intact: the number of bytes [`Self::write_to`]
    /// writes for such a package.
    ///
    /// Returns `None` exactly when [`Self::exact_source_sha256`] does. It reads
    /// the length without taking a handle to the archive or hashing it, so a
    /// caller can check a bound before asking for the digest.
    #[must_use]
    pub fn exact_source_len(&self) -> Option<usize> {
        self.exact_source().map(<[u8]>::len)
    }

    /// Whether two package snapshots retain the same immutable source archive.
    ///
    /// This is a preservation identity proof, not a content comparison: the
    /// archive allocation is immutable and every package clone or source
    /// descendant points at the same bytes. Format owners can use it to prove
    /// equality for opaque ZIP members which are intentionally outside the
    /// OPC part graph, while still comparing the exposed graph separately.
    #[must_use]
    pub fn shares_source_archive(&self, other: &Self) -> bool {
        self.source_archive
            .as_ref()
            .zip(other.source_archive.as_ref())
            .is_some_and(|(left, right)| Arc::ptr_eq(left.bytes(), right.bytes()))
    }

    /// The read limits this package's owned source archive was admitted
    /// under, or `None` when the package retains no owned source archive.
    ///
    /// A caller that re-admits the bytes this package publishes, for example
    /// to decide something from those bytes alone, can hold them to the
    /// policy the package itself was opened under (change 0742).
    #[must_use]
    pub fn source_read_limits(&self) -> Option<SourceLimits> {
        self.source_archive
            .as_ref()
            .map(|_archive| self.source_limits)
    }

    /// Configure font embedding with one self-documenting policy.
    pub fn with_fonts(&mut self, policy: FontEmbedding) -> &mut Self {
        self.revoke_exact_source();
        self.save_options.fonts = policy;
        self
    }

    /// Open an OPC package from a file.
    ///
    /// # Arguments
    /// * `path` - Path to the package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new `OpcPackage` instance loaded with the package contents
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::{package::OpcPackage, SourceLimits};
    ///
    /// let pkg = OpcPackage::open("document.docx", SourceLimits::default()).unwrap();
    /// ```
    ///
    /// # Errors
    /// Returns an error if the file cannot be read, violates `limits`, or is not
    /// a valid OPC package.
    pub fn open<P: AsRef<Path>>(path: P, limits: SourceLimits) -> Result<Self> {
        let data = read_owned_path_with_limits(path, limits)?;
        Self::from_owned_bytes_with_limits(data, limits)
    }

    /// Load an OPC package from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    ///
    /// # Errors
    /// Returns an error if the archive cannot be read, violates `limits`, or is
    /// not a valid OPC package.
    pub fn from_reader<R: Read>(reader: R, limits: SourceLimits) -> Result<Self> {
        let data = read_limited(reader, limits)?;
        Self::from_owned_bytes_with_limits(data, limits)
    }

    /// Move an owned ZIP archive into the package reader under explicit limits.
    ///
    /// This avoids copying the archive buffer before parts are decompressed.
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC package.
    pub fn from_vec(data: Vec<u8>, limits: SourceLimits) -> Result<Self> {
        Self::from_owned_bytes_with_limits(data, limits)
    }

    /// Load an owned OPC package while opportunistically reusing matching
    /// payload allocations from an already opened package.
    ///
    /// The new archive still goes through the complete physical-package and
    /// package-reader validation and decompression path.  For each part, the
    /// donor payload is selected only when its content type and bytes match
    /// the newly decoded payload, agrees with the donor's visible bytes,
    /// and its vector capacity is no larger. A
    /// failed comparison simply keeps the newly decoded allocation, so this
    /// optimization never changes package semantics.  Donor metadata,
    /// relationship state, save options, and source authorization are not
    /// carried into the returned package.
    ///
    /// # Errors
    ///
    /// Returns an error if the new archive violates `limits` or is not a
    /// valid OPC package.  The donor is used only after those checks and does
    /// not relax any read or allocation limit.
    pub fn from_vec_reusing_payloads(
        data: Vec<u8>,
        limits: SourceLimits,
        donor: &Self,
    ) -> Result<Self> {
        Self::from_shared_vec_reusing_payloads(Arc::new(data), limits, donor)
    }

    /// [`Self::from_vec_reusing_payloads`] over an archive allocation the
    /// caller keeps a handle to, instead of one it moves in.
    ///
    /// The package becomes a further owner of `data` rather than a copy of
    /// it: it reads, validates and decodes exactly as the owned ingress does,
    /// retains `data` as its exact owned source, and republishes it verbatim
    /// while that authorization is intact. Sharing cannot change what the
    /// package reads or publishes: a `Vec<u8>` behind an `Arc` held by more
    /// than one owner can be mutated in place by none of them, and no path of
    /// this crate mutates a retained source archive (change 0751).
    ///
    /// # Errors
    ///
    /// Returns an error if the archive violates `limits` or is not a valid
    /// OPC package. The donor is used only after those checks and does not
    /// relax any read or allocation limit.
    pub fn from_shared_vec_reusing_payloads(
        data: Arc<Vec<u8>>,
        limits: SourceLimits,
        donor: &Self,
    ) -> Result<Self> {
        let mut package = {
            let phys_reader = PhysPkgReader::new(data.as_slice(), limits)?;
            let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
            Self::unmarshal_with_payload_donor(pkg_reader, Some(donor))?
        };
        package.authorize_shared_owned_source(data, limits);
        Ok(package)
    }

    /// [`Self::from_vec`] over an archive allocation the caller
    /// keeps a handle to, instead of one it moves in.
    ///
    /// Parts are decoded on first access from `data`, exactly as for the owned
    /// ingress, and the package retains `data` as its exact owned source
    /// rather than a copy of it (see
    /// [`Self::from_shared_vec_reusing_payloads`] for why sharing is safe).
    ///
    /// # Errors
    ///
    /// Returns an error if the archive violates `limits` or is not a valid
    /// OPC package.
    pub fn from_shared_vec(data: Arc<Vec<u8>>, limits: SourceLimits) -> Result<Self> {
        Self::from_shared_bytes_with_limits(data, limits)
    }

    /// Build a deferred owned package from a physical reader that has already
    /// admitted the exact `Arc<Vec<u8>>` source under its retained policy.
    ///
    /// This is the handoff for callers that must perform a borrowed physical
    /// preflight before taking the package ownership path.  The reader is
    /// consumed by the deferred package reader, so its central directory is
    /// not parsed again.  The source allocation is shared, never copied, and
    /// the reader's policy remains the package's policy.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::SourceArtifactMismatch`] if `reader` was borrowed
    /// from another allocation, or the same errors as the ordinary deferred
    /// shared ingress for package-structure admission.
    pub fn from_shared_vec_with_admitted_reader(
        data: Arc<Vec<u8>>,
        reader: PhysPkgReader<'_>,
    ) -> Result<Self> {
        if !reader.source_matches(data.as_slice()) {
            return Err(OpcError::SourceArtifactMismatch {
                artifact: "shared OPC source",
                field: "physical reader bytes",
            });
        }
        let limits = reader.limits();
        let mut package = {
            let pkg_reader = PackageReader::from_phys_reader_deferred(&reader, &data)?;
            Self::unmarshal(pkg_reader)?
        };
        package.authorize_shared_owned_source(data, limits);
        Ok(package)
    }

    /// Moves an owned ZIP archive into an explicitly scheduled eager open.
    ///
    /// This additive advanced API retains exact owned-source authorization on
    /// success. Ordinary constructors remain serial and do not create an
    /// execution session.
    ///
    /// # Errors
    ///
    /// Returns a typed OPC, execution, or local-session error when opening
    /// cannot complete. Cancellation discards the incomplete package.
    pub fn from_vec_with_execution(
        data: Vec<u8>,
        limits: SourceLimits,
        execution: &OpenSession,
    ) -> Result<Self> {
        execution.from_vec(data, limits)
    }

    /// Load an OPC package from a byte slice.
    ///
    /// # Arguments
    /// * `data` - The ZIP archive data as a byte slice
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC
    /// package.
    pub fn from_bytes(data: &[u8], limits: SourceLimits) -> Result<Self> {
        let phys_reader = PhysPkgReader::new(data, limits)?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Loads a borrowed ZIP archive through an explicitly scheduled eager open.
    ///
    /// Ordinary constructors remain serial and do not create an execution
    /// session.
    ///
    /// # Errors
    ///
    /// Returns a typed OPC, execution, or local-session error when opening
    /// cannot complete. Cancellation discards the incomplete package.
    pub fn from_bytes_with_execution(
        data: &[u8],
        limits: SourceLimits,
        execution: &OpenSession,
    ) -> Result<Self> {
        execution.from_bytes(data, limits)
    }

    /// Unmarshal a package from a package reader.
    ///
    /// This is the main deserialization logic that converts serialized parts
    /// and relationships into the in-memory object graph.
    ///
    /// Optimized to minimize clones by consuming the package reader and moving data.
    pub(crate) fn unmarshal(pkg_reader: PackageReader) -> Result<Self> {
        Self::unmarshal_with_payload_donor(pkg_reader, None)
    }

    fn unmarshal_with_payload_donor(
        mut pkg_reader: PackageReader,
        donor: Option<&Self>,
    ) -> Result<Self> {
        let mut package = Self::new();

        package.read_limits = pkg_reader.read_limits();

        // Get ownership of package relationships, parts, and non-part members
        let pkg_srels = pkg_reader.take_pkg_srels();
        let mut source_relationships = pkg_reader.take_source_relationships();
        let (source_content_types_xml, source_content_types) =
            pkg_reader.take_source_content_types();
        let sparts = pkg_reader.take_sparts();
        package.non_part_members = pkg_reader.take_non_part_members();
        package.tolerance_report = pkg_reader.take_tolerance_report();
        package.source_content_types_xml = Some(source_content_types_xml);
        package.source_content_types = Some(Arc::new(source_content_types));

        // Pre-allocate with known capacity to avoid reallocations
        let mut parts_map: HashMap<PackURI, Box<dyn Part + Send + Sync>> = HashMap::new();
        parts_map
            .try_reserve(sparts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC package parts",
                source,
            })?;
        let mut source_xml_parts = HashMap::new();
        let mut part_bytes = 0u64;

        // Create all parts - move data instead of cloning
        for spart in sparts {
            let payload_len = u64::try_from(spart.payload.payload_len()?)
                .map_err(|_| OpcError::ZipError("OPC part payload length overflows u64".into()))?;
            part_bytes = part_bytes.checked_add(payload_len).ok_or_else(|| {
                OpcError::ZipError("OPC aggregate part payload length overflows u64".into())
            })?;
            let partname = spart.partname.clone(); // Need to clone partname for the HashMap key
            // Donation compares payloads, so it applies only to a reader that
            // already materialized them; a deferred payload is left alone.
            let donated = match (donor, spart.payload.decoded()) {
                (Some(donor), Some(read_blob)) => donor
                    .parts
                    .get(&partname)
                    .filter(|donor_part| donor_part.content_type() == spart.content_type.as_str())
                    .and_then(|donor_part| {
                        let blob = donor_part.blob_arc();
                        let visible = donor_part.blob();
                        // Built-in parts take the pointer fast path. A custom
                        // part cannot donate storage inconsistent with its blob.
                        (std::ptr::eq(blob.as_slice(), visible) || blob.as_slice() == visible)
                            .then_some(blob)
                    })
                    .filter(|donor_blob| {
                        donor_blob.capacity() <= read_blob.capacity()
                            && donor_blob.as_slice() == read_blob.as_slice()
                    })
                    .map(PartPayload::ready),
                _ => None,
            };
            let payload = donated.unwrap_or(spart.payload);
            let mut part = PartFactory::load_payload(
                spart.partname,     // Move
                spart.content_type, // Move
                payload,            // Move the selected payload storage
            )?;

            // Reserve the complete incoming relationship collection before
            // moving serialized edges into the part. This keeps the eager
            // unmarshal boundary fallible instead of allowing the map to
            // grow during insertion.
            part.rels_mut().try_reserve(spart.srels.len())?;

            // Load part relationships
            for srel in spart.srels {
                part.rels_mut().try_add_source_relationship(
                    srel.reltype,    // Move
                    srel.target_ref, // Move
                    srel.r_id,       // Move
                    srel.target_mode,
                )?;
            }

            if let Some(bytes) = source_relationships.remove(&partname) {
                package.retain_relationships_xml(partname.clone(), bytes, part.rels())?;
            }

            if xml_minifier::audit::package::is_xml_part(partname.as_str(), part.content_type()) {
                source_xml_parts
                    .try_reserve(1)
                    .map_err(|source| OpcError::Allocation {
                        resource: "OPC source-preserved XML parts",
                        source,
                    })?;
                source_xml_parts.insert(partname.clone(), part.payload_handle().payload().clone());
            }

            parts_map.insert(partname, part);
        }

        // Load package relationships - move instead of clone
        package.rels.try_reserve(pkg_srels.len())?;
        for srel in pkg_srels {
            package.rels.try_add_source_relationship(
                srel.reltype,    // Move
                srel.target_ref, // Move
                srel.r_id,       // Move
                srel.target_mode,
            )?;
        }

        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        if let Some(bytes) = source_relationships.remove(&package_uri) {
            let binding = RelationshipBinding::from_relationships(&package.rels)?;
            package
                .source_relationships_xml
                .try_reserve(1)
                .map_err(|source| OpcError::Allocation {
                    resource: "OPC relationship XML provenance",
                    source,
                })?;
            package.source_relationships_xml.insert(
                package_uri,
                Arc::new(SourceRelationshipsXml { bytes, binding }),
            );
        }

        package.parts = parts_map;
        package.part_bytes = Some(part_bytes);
        package.source_xml_parts = source_xml_parts;
        package.source_ingress = true;
        package.signature_graph_tracked = package.is_signed();
        package.signature_policy_required = package.signature_graph_tracked;
        Ok(package)
    }

    pub(crate) fn retain_relationships_xml(
        &mut self,
        owner: PackURI,
        bytes: Arc<Vec<u8>>,
        relationships: &Relationships,
    ) -> Result<()> {
        let binding = RelationshipBinding::from_relationships(relationships)?;
        self.source_relationships_xml
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC relationship XML provenance",
                source,
            })?;
        self.source_relationships_xml
            .insert(owner, Arc::new(SourceRelationshipsXml { bytes, binding }));
        Ok(())
    }

    /// Whether this XML part still holds the very allocation it was decoded
    /// from, and is therefore the source's own bytes rather than a caller's.
    ///
    /// This is a provenance proof, not a byte comparison. Ingress retains the
    /// `Arc` each XML part was loaded with
    /// ([`Self::unmarshal`] and [`Self::try_add_source_part`]); every
    /// documented route that changes a payload installs a different
    /// allocation, because [`Part::set_blob`] and [`Part::set_blob_shared`]
    /// replace the part's `Arc` rather than mutating through it, and
    /// [`Self::add_part`] and [`Self::remove_part`] drop the retained entry
    /// outright. So the entry survives exactly while the part is untouched:
    /// change 0647's "the map changed iff the capture was dropped" invariant,
    /// applied to payloads instead of relationship captures, and the same
    /// proof shape change 0593 gave part relationships.
    ///
    /// A caller that writes the source's own bytes back through `set_blob`
    /// stops being an original here, deliberately. The writer must audit the
    /// bytes a caller supplied, and bytes that merely *equal* the source's are
    /// indistinguishable from authored ones by inspection, so an equality test
    /// would let a mutation escape the audit. Change 0665 replaced that
    /// equality test with this proof; the shared `source_blob_retained`
    /// publication predicate keeps its byte comparison, because an equal
    /// payload publishes equal bytes either way.
    ///
    /// Publication does not decide compactness from this signal — no member is
    /// refused for spelling any more (change 0665) — only whether the member
    /// is parsed again before it is republished.
    pub(crate) fn holds_original_source_xml(&self, part: &dyn Part) -> bool {
        self.source_xml_parts
            .get(part.partname())
            .is_some_and(|source| {
                if part.decoded_blob().is_none() {
                    // A still-deferred original member has never exposed its
                    // payload for replacement; ADR 0030 permits raw passthrough.
                    return true;
                }
                // Keep 0665's allocation provenance. Equal replacement bytes
                // must still pass publication audit, even after lazy ingress.
                source
                    .decoded()
                    .is_some_and(|source| Arc::ptr_eq(source, &part.blob_arc()))
            })
    }

    /// Capture bounded XML provenance from this owned package. Noncompact XML
    /// must match retained source bytes; new XML must pass the authored audit.
    pub fn source_xml_part(&self, name: &PackURI) -> Result<crate::OwnedXmlPart> {
        let part = self.get_part(name)?;
        if !self.holds_original_source_xml(part)
            && crate::authored_xml_requires_source_proof(name, part.content_type(), part.blob())?
        {
            return Err(OpcError::XmlError(
                "noncompact owned XML has no retained source provenance".into(),
            ));
        }
        let bytes = part.blob_arc();
        if bytes.as_slice() != part.blob() {
            return Err(OpcError::XmlError(
                "part storage differs from its visible XML".into(),
            ));
        }
        crate::OwnedXmlPart::capture_with_limits(
            name.clone(),
            part.content_type().into(),
            bytes,
            self.read_limits,
        )
    }

    /// Replace one exact expected XML source with a validated provenance token.
    /// This preserves untouched lexical XML without weakening authored output
    /// validation. A stale source or content type fails before publication.
    pub fn try_replace_owned_xml_part(
        &mut self,
        expected: &[u8],
        replacement: crate::OwnedXmlPart,
    ) -> Result<()> {
        if self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        let current = self.get_part(&replacement.name)?;
        let canonical_name = current.partname().clone();
        // A deferred payload reports its own refusal rather than comparing as
        // an empty member (ADR 0030).
        current.ensure_payload()?;
        if current.content_type() != replacement.content_type || current.blob() != expected {
            return Err(OpcError::XmlError(
                "stale owned XML part replacement".into(),
            ));
        }
        // A token may have been captured under a broader producer profile.
        // Re-admit it under this package's retained profile before installing
        // provenance; content equality alone must never widen the recipient.
        replacement.preflight_under(self.read_limits)?;
        // These bytes are recorded as the part's source provenance below, and
        // provenance exempts a payload from the writer's publication audit
        // (change 0665). They came from the caller, so audit them now, before
        // any mutation, and hand the part the proof (change 0754).
        let verified = verified_source_payload(&canonical_name, Arc::clone(&replacement.bytes))?;
        self.source_xml_parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML provenance",
                source,
            })?;
        let part = self.get_part_mut(&canonical_name)?;
        part.set_blob_verified(verified);
        let payload = part.payload_handle().payload().clone();
        self.source_xml_parts.insert(canonical_name, payload);
        Ok(())
    }

    /// Replace one existing XML part with a validated source-preserving
    /// payload while checking the caller's expected current bytes.
    ///
    /// This is the byte-oriented companion to
    /// [`Self::try_replace_owned_xml_part`].  Format-owned callers use it when
    /// their source-backed resource already retains the replacement allocation
    /// but does not expose the OPC token type.  The expected bytes remain the
    /// stale-source guard; the replacement is validated with the current
    /// part's content type before its source provenance is installed.
    pub fn try_replace_owned_xml_part_bytes(
        &mut self,
        partname: &PackURI,
        expected: &[u8],
        replacement: Arc<Vec<u8>>,
    ) -> Result<()> {
        let current = self.get_part(partname)?;
        let canonical_name = current.partname().clone();
        let content_type = current.content_type().to_owned();
        let token = crate::OwnedXmlPart::capture_with_limits(
            canonical_name,
            content_type,
            replacement,
            self.read_limits,
        )?;
        self.try_replace_owned_xml_part(expected, token)
    }

    /// Add previously validated source XML, for exact restoration or transfer.
    /// Relationships remain the responsibility of the format-owned graph edit.
    pub fn try_add_owned_xml_part(&mut self, source: crate::OwnedXmlPart) -> Result<()> {
        if self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        source.preflight_under(self.read_limits)?;
        let mut part = crate::BlobPart::new_shared(source.name, source.content_type, source.bytes);
        // The part is recorded as source provenance below, and provenance
        // exempts a payload from the writer's publication audit (change
        // 0665). Its bytes came from the caller, so audit them now, before
        // any mutation, and hand the part the proof (change 0754).
        if xml_minifier::audit::package::is_xml_part(part.partname().as_str(), part.content_type())
        {
            let verified = verified_source_payload(part.partname(), part.blob_arc())?;
            part.set_blob_verified(verified);
        }
        self.try_add_source_part(Box::new(part))
    }

    /// Add a validated source-preserving XML part from its shared payload.
    ///
    /// This is the byte-oriented companion to [`Self::try_add_owned_xml_part`]
    /// for format-owned resources whose source token is represented by the
    /// retained allocation itself.
    pub fn try_add_owned_xml_part_bytes(
        &mut self,
        name: PackURI,
        content_type: String,
        bytes: Arc<Vec<u8>>,
    ) -> Result<()> {
        let source =
            crate::OwnedXmlPart::capture_with_limits(name, content_type, bytes, self.read_limits)?;
        self.try_add_owned_xml_part(source)
    }

    /// Add caller-authored XML while retaining its exact bytes as publication
    /// provenance.
    ///
    /// This is for format-owned opaque XML whose semantics have already been
    /// validated by the format codec. It performs the ordinary bounded XML
    /// publication audit, but does not issue a [`crate::OwnedXmlPart`] source-splice
    /// token: source-splice tokens require literal namespace declarations so
    /// expanded-name edits cannot be ambiguous, while a whole-part authored
    /// replacement preserves the caller's namespace spelling byte-for-byte.
    /// The shared payload is retained without copying and is published through
    /// the same source-preserving writer as an exact source part. The retained
    /// allocation is publication provenance only; this does not create an
    /// [`crate::OwnedXmlPart`] or [`crate::SourceXmlPart`] range-splice token.
    pub fn try_add_authored_xml_part_bytes(
        &mut self,
        name: PackURI,
        content_type: String,
        bytes: Arc<Vec<u8>>,
        limits: SourceLimits,
    ) -> Result<()> {
        if self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        if !xml_minifier::audit::package::is_xml_part(name.as_str(), &content_type) {
            return Err(OpcError::XmlError(
                "authored XML part content type is not XML".to_owned(),
            ));
        }
        let limits = self.read_limits.min(limits);
        let mut part = crate::BlobPart::new_shared(name, content_type, Arc::clone(&bytes));
        let source_xml = self.preflight_authored_source_part_addition(&part, limits)?;
        crate::OwnedXmlPart::check_capture_size(part.partname(), bytes.len(), limits)?;
        let verified = verified_source_payload_with_source_limits(
            part.partname(),
            Arc::clone(&bytes),
            limits,
        )?;
        crate::xml_splice::validate_authored_xml_bytes(part.partname(), &bytes, limits)?;
        part.set_blob_verified(verified);
        self.install_prevalidated_source_part(Box::new(part), source_xml)
    }

    /// Replace one exact XML payload with caller-authored bytes while
    /// retaining those bytes as publication provenance.
    ///
    /// The expected payload is an exact stale-source guard. The replacement
    /// is admitted by the bounded publication audit rather than the stricter
    /// source-splice namespace proof, because the format owner is replacing
    /// the whole opaque XML member and does not splice its internals. Equal
    /// replacement bytes are an audited no-op, including for a signed package;
    /// non-no-op failures leave package and signature state untouched.
    pub fn try_replace_authored_xml_part_bytes(
        &mut self,
        partname: &PackURI,
        expected: &[u8],
        replacement: Arc<Vec<u8>>,
        limits: SourceLimits,
    ) -> Result<()> {
        let current = self.get_part(partname)?;
        let canonical_name = current.partname().clone();
        let content_type = current.content_type().to_owned();
        if !xml_minifier::audit::package::is_xml_part(canonical_name.as_str(), &content_type) {
            return Err(OpcError::XmlError(
                "authored XML part content type is not XML".to_owned(),
            ));
        }
        current.ensure_payload()?;
        if current.blob() != expected {
            return Err(OpcError::XmlError(
                "stale authored XML part replacement".to_owned(),
            ));
        }
        let limits = self.read_limits.min(limits);
        let unchanged = current.blob() == replacement.as_slice();
        if self.requires_signature_edit_policy() && !unchanged {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        crate::OwnedXmlPart::check_capture_size(&canonical_name, replacement.len(), limits)?;
        self.check_authored_xml_replacement_limits(&canonical_name, replacement.len(), limits)?;
        let verified = verified_source_payload_with_source_limits(
            &canonical_name,
            Arc::clone(&replacement),
            limits,
        )?;
        crate::xml_splice::validate_authored_xml_bytes(&canonical_name, &replacement, limits)?;
        if unchanged {
            return Ok(());
        }
        self.source_xml_parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "owned authored XML provenance",
                source,
            })?;
        let part = self.get_part_mut(&canonical_name)?;
        part.set_blob_verified(verified);
        let payload = part.payload_handle().payload().clone();
        self.source_xml_parts.insert(canonical_name, payload);
        Ok(())
    }

    /// Get a reference to the main document part.
    ///
    /// For Word documents, this is the document.xml part.
    /// For Excel, the workbook.xml part.
    /// For `PowerPoint`, the presentation.xml part.
    ///
    /// # Errors
    /// Returns an error if the package has no main-document relationship, has
    /// more than one, the relationship is external, or the target part is missing.
    pub fn main_document_part(&self) -> Result<&dyn Part> {
        let mut matching = self.rels.iter().filter(|relationship| {
            matches!(
                relationship.reltype(),
                relationship_type::OFFICE_DOCUMENT | relationship_type::STRICT_OFFICE_DOCUMENT
            )
        });
        let rel = matching.next().ok_or_else(|| {
            OpcError::InvalidRelationship("main-document relationship is missing".to_string())
        })?;
        if matching.next().is_some() {
            return Err(OpcError::InvalidRelationship(
                "package has multiple main-document relationships".to_string(),
            ));
        }
        if rel.is_external() {
            return Err(OpcError::InvalidRelationship(
                "main-document relationship cannot be external".to_string(),
            ));
        }
        if let Some(memo) = self.main_partname.get()
            && memo.resolves(rel)
        {
            return self.get_part(&memo.partname);
        }
        let partname = rel.target_partname()?;
        let part = self.get_part(&partname)?;
        let _first = self.main_partname.set(Arc::new(MainPartname {
            relationship: rel.clone(),
            partname,
        }));
        Ok(part)
    }

    /// Get a part by its partname.
    ///
    /// # Arguments
    /// * `partname` - The `PackURI` of the part to retrieve
    ///
    /// # Errors
    /// Returns `OpcError::PartNotFound` if no part with `partname` exists.
    pub fn get_part(&self, partname: &PackURI) -> Result<&dyn Part> {
        let part = if let Some(part) = self.parts.get(partname) {
            let part_ref: &dyn Part = &**part;
            part_ref
        } else {
            self.find_case_insensitive(partname)
                .map(|(_, part)| part)
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?
        };
        // No part leaves this crate before its payload is decoded, so a
        // caller never receives a part whose bytes are still compressed and
        // never has a decode refusal swallowed by an infallible accessor.
        part.ensure_payload()?;
        Ok(part)
    }

    /// Look up a part's name and metadata without decoding its payload.
    ///
    /// Names resolve exactly as in [`Self::get_part`]: exact match first,
    /// then ASCII case-insensitive. `None` therefore means exactly "no such
    /// part". Do not use `get_part(..).is_err()` as an absence test: an
    /// owned-source package decodes a payload on first access (ADR 0030), so
    /// that error also covers a present part whose payload fails to decode.
    /// The returned [`PartMetadata`] gives the stored name, content type and
    /// relationships, and cannot reach the payload.
    #[must_use]
    pub fn part_metadata(&self, partname: &PackURI) -> Option<PartMetadata<'_>> {
        let part: &dyn Part = match self.parts.get(partname) {
            Some(part) => &**part,
            None => self.find_case_insensitive(partname)?.1,
        };
        Some(PartMetadata::new(part))
    }

    /// Locate a part whose name matches `partname` ignoring ASCII case.
    ///
    /// OPC compares part names case-insensitively, which is why a package
    /// containing two names differing only by case is rejected as ambiguous
    /// when it is read. Because that ambiguity cannot survive loading, this
    /// fallback can match at most one part, and it is what lets a package whose
    /// writer stored `/xl/sharedstrings.xml` still resolve a relationship
    /// targeting `sharedStrings.xml` — without it those parts are simply
    /// unreachable and their content silently reads as absent.
    ///
    /// The exact lookup above is the fast path; this linear scan runs only on a
    /// miss, and the part count is already bounded when the package is read.
    fn find_case_insensitive(&self, partname: &PackURI) -> Option<(&PackURI, &dyn Part)> {
        let wanted = partname.as_str();
        self.parts
            .iter()
            .find(|(name, _)| name.as_str().eq_ignore_ascii_case(wanted))
            .map(|(name, part)| {
                let part_ref: &dyn Part = &**part;
                (name, part_ref)
            })
    }

    /// Get a mutable reference to a part by its partname.
    ///
    /// # Errors
    /// Returns `OpcError::PartNotFound` if no part with `partname` exists.
    pub fn get_part_mut(&mut self, partname: &PackURI) -> Result<&mut dyn Part> {
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        // A mutable Part exposes its relationship collection, so retain the
        // signature audit even when the caller edits that collection directly.
        self.signature_graph_tracked = true;
        // Resolve the stored key first so the borrow of `self.parts` ends before
        // the mutable lookup; see `find_case_insensitive` for why this matches.
        let key = if self.parts.contains_key(partname) {
            partname.clone()
        } else {
            match self.find_case_insensitive(partname) {
                Some((name, _)) => name.clone(),
                None => return Err(OpcError::PartNotFound(partname.to_string())),
            }
        };
        // Decode through the shared borrow before handing out the mutable
        // one. A caller that replaces the payload must still observe the
        // refusal a failed decode records, and a caller that reads it must
        // never see an undecoded part.
        match self.parts.get(&key) {
            Some(part) => part.ensure_payload()?,
            None => return Err(OpcError::PartNotFound(partname.to_string())),
        }
        // The mutable trait object can replace bytes or relationships through
        // methods this owner cannot observe. Publication callers refresh the
        // aggregate from their existing capture walk.
        self.part_bytes = None;
        self.parts
            .get_mut(&key)
            .map(|b| {
                let part: &mut dyn Part = &mut **b;
                part
            })
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Replace the payloads of existing parts as one prevalidated raw edit.
    ///
    /// Every requested part is resolved before the first replacement. Exact
    /// names win, then OPC's ASCII-case-insensitive fallback is used. A
    /// repeated target, including one written through a case alias, is
    /// rejected before mutation. This makes the duplicate policy explicit and
    /// prevents two callers from silently disagreeing about one payload.
    ///
    /// The empty batch is a no-op and keeps exact-source authorization intact.
    /// Nonempty batches require a package containing only built-in
    /// [`BlobPart`](crate::BlobPart) and [`XmlPart`](crate::XmlPart) values.
    /// Custom parts are refused because their setters cannot guarantee atomic
    /// failure. All target payloads are decoded before mutation, and any
    /// returned error leaves payloads and source authorization unchanged.
    ///
    /// Exact-source authorization is revoked and signature policy recorded
    /// once before the setters run. Replacements use [`Part::set_blob`], so
    /// caller-authored bytes receive no retained XML source provenance.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::PartNotFound`] for an unknown target,
    /// [`OpcError::DuplicatePartPayload`] for a repeated resolved target,
    /// the selected part's deferred decode error, or
    /// [`OpcError::UnsupportedCustomPartPayload`] when any custom part exists.
    /// Allocation failures are returned as [`OpcError::Allocation`].
    pub fn replace_part_payloads(&mut self, replacements: Vec<(PackURI, Vec<u8>)>) -> Result<()> {
        if replacements.is_empty() {
            return Ok(());
        }

        let mut prepared = Vec::new();
        prepared
            .try_reserve_exact(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch payload replacements",
                source,
            })?;
        let mut seen = HashSet::new();
        seen.try_reserve(replacements.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch payload target names",
                source,
            })?;

        // Resolve and decode every target before changing package state. The
        // decoded payload can still be deferred in an owned source, so this
        // pass is also the batch's all-or-nothing decode admission boundary.
        for (requested, payload) in replacements {
            let partname = if self.parts.contains_key(&requested) {
                requested.clone()
            } else {
                self.find_case_insensitive(&requested)
                    .map(|(stored, _)| stored.clone())
                    .ok_or_else(|| OpcError::PartNotFound(requested.to_string()))?
            };
            if !seen.insert(partname.clone()) {
                return Err(OpcError::DuplicatePartPayload(partname.to_string()));
            }
            let part = self
                .parts
                .get(&partname)
                .map(|part| &**part)
                .ok_or_else(|| OpcError::PartNotFound(requested.to_string()))?;
            part.ensure_payload()?;
            prepared.push((partname, payload));
        }

        if !self.holds_only_built_in_parts() {
            return Err(OpcError::UnsupportedCustomPartPayload);
        }
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        self.part_bytes = None;
        for (partname, payload) in prepared {
            // Preflight resolved every key and no map mutation intervenes.
            // Built-in setters cannot change relationships or the part map.
            self.parts
                .get_mut(&partname)
                .expect("preflight resolved the unchanged part map")
                .set_blob(payload);
        }
        Ok(())
    }

    /// Get a part by relationship type from the package level.
    ///
    /// # Arguments
    /// * `reltype` - The relationship type URI
    ///
    /// # Errors
    /// Returns an error if no relationship of `reltype` exists or the target
    /// part is missing.
    pub fn part_by_reltype(&self, reltype: &str) -> Result<&dyn Part> {
        let rel = self.rels.part_with_reltype(reltype)?;
        if let Some(memo) = self.main_partname.get()
            && memo.resolves(rel)
        {
            return self.get_part(&memo.partname);
        }
        let partname = rel.target_partname()?;
        let part = self.get_part(&partname)?;
        let _first = self.main_partname.set(Arc::new(MainPartname {
            relationship: rel.clone(),
            partname,
        }));
        Ok(part)
    }

    /// Add a new part to the package.
    ///
    /// # Arguments
    /// * `part` - The part to add
    pub fn add_part(&mut self, part: Box<dyn Part + Send + Sync>) {
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.signature_graph_tracked = true;
        self.mark_content_types_invalid_if_uncovered(&*part);
        let partname = part.partname().clone();
        self.source_xml_parts.remove(&partname);
        self.source_relationships_xml.remove(&partname);
        self.part_bytes = None;
        self.parts.insert(partname, part);
    }

    /// Try to add a part without replacing an existing or ambiguous part name.
    ///
    /// # Errors
    /// Returns an error if the part's partname duplicates or conflicts with an
    /// existing part name; the existing part is left untouched.
    pub fn try_add_part(&mut self, part: Box<dyn Part + Send + Sync>) -> Result<()> {
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.signature_graph_tracked = true;
        let partname = part.partname().clone();
        self.read_limits.check(
            crate::limits::SourceLimitKind::ArchiveMemberNameBytes,
            partname.membername().len() as u64,
            self.read_limits.max_archive_member_name_bytes(),
        )?;
        self.validate_new_part_name(&partname)?;
        self.mark_content_types_invalid_if_uncovered(&*part);
        self.parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC package parts",
                source,
            })?;
        self.source_relationships_xml.remove(&partname);
        self.part_bytes = None;
        self.parts.insert(partname, part);
        Ok(())
    }

    /// Add a prevalidated batch while retaining exact source content-types and
    /// relationship tokens.  Format-owned callers use this when a graph edit
    /// must publish several new targets atomically; the ordinary single-part
    /// API intentionally keeps its existing authored-source behavior.
    pub fn try_add_parts_with_source_tokens(
        &mut self,
        expected_content_types: &[u8],
        replacement_content_types: &OwnedContentTypes,
        expected_relationships: &OwnedRelationships,
        replacement_relationships: &OwnedRelationships,
        mut parts: Vec<Box<dyn Part + Send + Sync>>,
    ) -> Result<()> {
        let limits = self.read_limits;
        let current_content_types = self.source_content_types(limits)?;
        if current_content_types.bytes() != expected_content_types {
            return Err(OpcError::InvalidContentTypesManifest(
                "stale content-types source replacement".to_owned(),
            ));
        }
        let current_relationships =
            self.source_relationships(expected_relationships.owner(), self.read_limits)?;
        if current_relationships != *expected_relationships {
            return Err(OpcError::InvalidRelationship(
                "stale relationship replacement".to_owned(),
            ));
        }
        if replacement_relationships.owner() != expected_relationships.owner() {
            return Err(OpcError::InvalidRelationship(
                "relationship replacement has a different owner".to_owned(),
            ));
        }
        if replacement_content_types.xml.name.as_str() != CONTENT_TYPES_URI
            || replacement_content_types.xml.content_type != crate::constants::content_type::XML
        {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types replacement has an invalid owner or content type".to_owned(),
            ));
        }
        if self.is_signed() || self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        // A source token's retained binding proves only its producer policy.
        // Re-admit both replacement members under this recipient before any
        // graph validation or package mutation.  The XML preflight covers
        // generic member/part ceilings; the typed parses cover their own
        // mapping, event and relationship dimensions.
        replacement_content_types.xml.preflight_under(limits)?;
        let replacement_content_types_binding = Arc::new(ContentTypeMap::from_xml(
            replacement_content_types.bytes(),
            limits,
        )?);
        if replacement_content_types_binding.as_ref() != replacement_content_types.binding.as_ref()
        {
            return Err(OpcError::InvalidContentTypesManifest(
                "content-types replacement binding does not match its XML".to_owned(),
            ));
        }
        replacement_relationships.preflight_under(limits)?;
        let replacement_relationships_parsed = PackageReader::parse_owned_relationships(
            replacement_relationships.bytes(),
            replacement_relationships.owner(),
            limits,
        )?;
        let replacement_relationship_binding =
            RelationshipBinding::from_relationships(&replacement_relationships_parsed)?;
        if !replacement_relationships.member_present()
            && !replacement_relationships_parsed.is_empty()
        {
            return Err(OpcError::InvalidRelationship(
                "absent relationship member has edges".to_owned(),
            ));
        }
        let mut new_names = Vec::new();
        new_names
            .try_reserve_exact(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch part names",
                source,
            })?;
        for part in &parts {
            let name = part.partname();
            self.validate_new_part_name(name)?;
            let declared = replacement_content_types_binding
                .lookup(name)
                .ok_or_else(|| OpcError::ContentTypeNotFound(name.to_string()))?;
            if declared.as_str() != part.content_type() {
                return Err(OpcError::InvalidContentTypesManifest(
                    "replacement content-types mapping does not match a new part".to_owned(),
                ));
            }
            new_names.push(name.clone());
        }
        new_names.sort_unstable_by(|left, right| {
            cmp_ascii_case_insensitive(left.as_str(), right.as_str())
        });
        for pair in new_names.windows(2) {
            if let Some(conflict) = pair[0].conflict_with(&pair[1]) {
                return Err(part_name_conflict_error(&pair[0], &pair[1], conflict));
            }
        }
        for part in self.parts.values() {
            let declared = replacement_content_types_binding
                .lookup(part.partname())
                .ok_or_else(|| OpcError::ContentTypeNotFound(part.partname().to_string()))?;
            if declared.as_str() != part.content_type() {
                return Err(OpcError::InvalidContentTypesManifest(
                    "replacement content-types mapping does not match an existing part".to_owned(),
                ));
            }
        }
        for relationship in replacement_relationships_parsed.iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            // Existence is a question about names: decoding the target here
            // would report a present part whose payload fails to decode as
            // missing (ADR 0030).
            let existing = self.part_metadata(&target).is_some()
                || new_names
                    .binary_search_by(|name| {
                        cmp_ascii_case_insensitive(name.as_str(), target.as_str())
                    })
                    .is_ok();
            if !existing {
                return Err(OpcError::InvalidRelationship(
                    "replacement relationship targets a missing part".to_owned(),
                ));
            }
        }
        self.parts
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch parts",
                source,
            })?;
        // New XML parts are recorded as source provenance below, which exempts
        // them from the writer's publication audit (change 0665). They came
        // from the caller, so audit each one now, before any mutation, and hand
        // it the proof (change 0754).
        let mut xml_count = 0usize;
        for part in &mut parts {
            if !xml_minifier::audit::package::is_xml_part(
                part.partname().as_str(),
                part.content_type(),
            ) {
                continue;
            }
            part.ensure_payload()?;
            let verified = verified_source_payload(part.partname(), part.blob_arc())?;
            part.set_blob_verified(verified);
            xml_count += 1;
        }
        // The two tokens become the retained manifest and the owner's `.rels`
        // source, which the writer publishes without its audit as well, so
        // caller bytes are audited here too. A token equal to what the package
        // already retains is that source, not caller XML, and stays exempt.
        if replacement_content_types.bytes() != current_content_types.bytes() {
            verified_metadata_token(
                &replacement_content_types.xml.name,
                Arc::clone(&replacement_content_types.xml.bytes),
                limits.max_content_types_bytes(),
            )?;
        }
        if replacement_relationships.member_present()
            && replacement_relationships.bytes() != current_relationships.bytes()
        {
            let member = replacement_relationships
                .owner()
                .rels_uri()
                .map_err(OpcError::InvalidPackUri)?;
            verified_metadata_token(
                &member,
                replacement_relationships.bytes_arc(),
                limits.max_relationship_xml_bytes(),
            )?;
        }
        self.source_xml_parts
            .try_reserve(xml_count)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch source XML parts",
                source,
            })?;
        self.source_relationships_xml
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC batch relationship provenance",
                source,
            })?;

        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.signature_graph_tracked = true;
        self.part_bytes = None;
        for part in parts {
            let name = part.partname().clone();
            let source_blob =
                if xml_minifier::audit::package::is_xml_part(name.as_str(), part.content_type()) {
                    Some(part.payload_handle().payload().clone())
                } else {
                    None
                };
            self.source_relationships_xml.remove(&name);
            self.parts.insert(name.clone(), part);
            if let Some(source_blob) = source_blob {
                self.source_xml_parts.insert(name, source_blob);
            }
        }
        if replacement_relationships.owner().as_str() == "/" {
            self.rels = replacement_relationships_parsed.clone();
        } else {
            let owner = self
                .parts
                .get_mut(replacement_relationships.owner())
                .ok_or_else(|| {
                    OpcError::PartNotFound(replacement_relationships.owner().to_string())
                })?;
            *owner.rels_mut() = replacement_relationships_parsed.clone();
        }
        if replacement_relationships.member_present() {
            self.source_relationships_xml.insert(
                replacement_relationships.owner().clone(),
                Arc::new(SourceRelationshipsXml {
                    bytes: replacement_relationships.bytes_arc(),
                    binding: replacement_relationship_binding,
                }),
            );
        } else {
            self.source_relationships_xml
                .remove(replacement_relationships.owner());
        }
        self.source_content_types_xml = Some(Arc::clone(&replacement_content_types.xml.bytes));
        self.source_content_types = Some(replacement_content_types_binding);
        Ok(())
    }

    /// Add a source-materialized part while retaining its exact XML bytes for
    /// the publication audit. Source-backed conversion uses this narrow seam
    /// so unchanged opaque XML remains publishable without reparsing or
    /// normalizing it. The source blob is shared with the inserted part and
    /// the metadata map grows through a fallible reservation. The retained
    /// bytes are publication provenance only; this path does not create an
    /// [`crate::OwnedXmlPart`] or [`crate::SourceXmlPart`] range-splice token.
    /// Keep this private source-materialization route cheap: its callers have
    /// already admitted their payloads, so it checks names and reservations
    /// without rescanning every existing payload for an aggregate total.
    pub(crate) fn try_add_source_part(&mut self, part: Box<dyn Part + Send + Sync>) -> Result<()> {
        let partname = part.partname().clone();
        self.validate_new_part_name(&partname)?;
        let source_xml =
            xml_minifier::audit::package::is_xml_part(partname.as_str(), part.content_type());
        self.install_prevalidated_source_part(part, source_xml)
    }

    fn preflight_authored_source_part_addition(
        &self,
        part: &dyn Part,
        limits: SourceLimits,
    ) -> Result<bool> {
        let partname = part.partname();
        let maximum_parts = u64::try_from(limits.max_parts()).unwrap_or(u64::MAX);
        limits.check(
            crate::SourceLimitKind::ArchiveMemberNameBytes,
            u64::try_from(partname.membername().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
        self.validate_new_part_name(partname)?;
        let payload_len = part.payload_len()?;
        let payload_bytes = u64::try_from(payload_len).map_err(|_| OpcError::Limit {
            kind: crate::SourceLimitKind::PartBytes,
            observed: u64::MAX,
            maximum: limits.max_part_bytes(),
        })?;
        limits.check(
            crate::SourceLimitKind::PartBytes,
            payload_bytes,
            limits.max_part_bytes(),
        )?;
        let next_parts = self.parts.len().checked_add(1).ok_or(OpcError::Limit {
            kind: crate::SourceLimitKind::Parts,
            observed: u64::MAX,
            maximum: maximum_parts,
        })?;
        limits.check(
            crate::SourceLimitKind::Parts,
            u64::try_from(next_parts).unwrap_or(u64::MAX),
            maximum_parts,
        )?;
        let mut total_bytes = payload_bytes;
        for existing in self.parts.values() {
            let bytes = u64::try_from(existing.payload_len()?).map_err(|_| OpcError::Limit {
                kind: crate::SourceLimitKind::TotalPartBytes,
                observed: u64::MAX,
                maximum: limits.max_total_part_bytes(),
            })?;
            total_bytes = total_bytes.checked_add(bytes).ok_or(OpcError::Limit {
                kind: crate::SourceLimitKind::TotalPartBytes,
                observed: u64::MAX,
                maximum: limits.max_total_part_bytes(),
            })?;
        }
        limits.check(
            crate::SourceLimitKind::TotalPartBytes,
            total_bytes,
            limits.max_total_part_bytes(),
        )?;
        Ok(xml_minifier::audit::package::is_xml_part(
            partname.as_str(),
            part.content_type(),
        ))
    }

    fn check_authored_xml_replacement_limits(
        &self,
        partname: &PackURI,
        replacement_len: usize,
        limits: SourceLimits,
    ) -> Result<()> {
        let replacement_bytes = u64::try_from(replacement_len).map_err(|_| OpcError::Limit {
            kind: crate::SourceLimitKind::PartBytes,
            observed: u64::MAX,
            maximum: limits.max_part_bytes(),
        })?;
        limits.check(
            crate::SourceLimitKind::PartBytes,
            replacement_bytes,
            limits.max_part_bytes(),
        )?;
        limits.check(
            crate::SourceLimitKind::Parts,
            u64::try_from(self.parts.len()).unwrap_or(u64::MAX),
            u64::try_from(limits.max_parts()).unwrap_or(u64::MAX),
        )?;
        let mut total_bytes = 0u64;
        for existing in self.parts.values() {
            let bytes = if existing
                .partname()
                .as_str()
                .eq_ignore_ascii_case(partname.as_str())
            {
                replacement_bytes
            } else {
                u64::try_from(existing.payload_len()?).map_err(|_| OpcError::Limit {
                    kind: crate::SourceLimitKind::TotalPartBytes,
                    observed: u64::MAX,
                    maximum: limits.max_total_part_bytes(),
                })?
            };
            total_bytes = total_bytes.checked_add(bytes).ok_or(OpcError::Limit {
                kind: crate::SourceLimitKind::TotalPartBytes,
                observed: u64::MAX,
                maximum: limits.max_total_part_bytes(),
            })?;
        }
        limits.check(
            crate::SourceLimitKind::TotalPartBytes,
            total_bytes,
            limits.max_total_part_bytes(),
        )
    }

    fn install_prevalidated_source_part(
        &mut self,
        part: Box<dyn Part + Send + Sync>,
        source_xml: bool,
    ) -> Result<()> {
        let partname = part.partname().clone();
        self.parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC source-materialized parts",
                source,
            })?;
        if source_xml {
            self.source_xml_parts
                .try_reserve(1)
                .map_err(|source| OpcError::Allocation {
                    resource: "OPC source-preserved XML parts",
                    source,
                })?;
        }
        let source_blob = source_xml.then(|| part.payload_handle());
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        self.content_types_source_state.reset();
        self.source_relationships_xml.remove(&partname);
        self.part_bytes = None;
        self.parts.insert(partname.clone(), part);
        if let Some(source_blob) = source_blob {
            self.source_xml_parts
                .insert(partname, source_blob.payload().clone());
        }
        Ok(())
    }

    /// Add a complete set of parts produced by source materialization.
    ///
    /// The ordinary insertion methods deliberately validate one name against
    /// the package map on every call.  A materialization already has all
    /// incoming parts in hand, so doing that would turn a package with `P`
    /// parts into a quadratic operation.  This private seam admits the whole
    /// batch through the same [`PartNameIndex`] used at source ingress, then
    /// reserves every destination map before changing package state.  The
    /// boolean paired with each part says whether its payload is retained as
    /// source XML provenance; authored replacements must pass `false`.
    ///
    /// `checkpoint` is called while the indexed preflight is running and
    /// after insertion.  Source-backed callers use it for their normal
    /// freshness and execution-context fence, retaining the same cancellation
    /// boundary as the per-part insertion loop.  A checkpoint that fails after
    /// insertion may leave this destination partially populated; this seam is
    /// crate-private and is used only with a disposable conversion-local
    /// package, which callers discard on every error.
    pub(crate) fn try_add_materialized_parts<F>(
        &mut self,
        parts: Vec<(Box<dyn Part + Send + Sync>, bool)>,
        mut checkpoint: F,
    ) -> Result<()>
    where
        F: FnMut() -> Result<()>,
    {
        if parts.is_empty() {
            return Ok(());
        }
        checkpoint()?;

        let capacity = self.parts.len().checked_add(parts.len()).ok_or_else(|| {
            OpcError::ZipError("OPC materialized part count overflows".to_owned())
        })?;
        let mut part_names = PartNameIndex::try_with_capacity(capacity)?;
        for name in self.parts.keys() {
            checkpoint()?;
            part_names.insert(name)?;
        }

        let mut source_xml_count = 0usize;
        for (part, retain_source) in &parts {
            checkpoint()?;
            part_names.insert(part.partname())?;
            if *retain_source
                && xml_minifier::audit::package::is_xml_part(
                    part.partname().as_str(),
                    part.content_type(),
                )
            {
                source_xml_count = source_xml_count.checked_add(1).ok_or_else(|| {
                    OpcError::ZipError("OPC source XML part count overflows".to_owned())
                })?;
            }
        }
        // The index has completed the full namespace admission. Drop its
        // temporary allocations before reserving the owning package maps so a
        // large eager conversion does not retain two indexes while inserting.
        drop(part_names);

        self.parts
            .try_reserve(parts.len())
            .map_err(|source| OpcError::Allocation {
                resource: "OPC materialized parts",
                source,
            })?;
        if source_xml_count != 0 {
            self.source_xml_parts
                .try_reserve(source_xml_count)
                .map_err(|source| OpcError::Allocation {
                    resource: "OPC materialized source-preserved XML parts",
                    source,
                })?;
        }
        checkpoint()?;

        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        self.part_bytes = None;
        for (part, retain_source) in parts {
            let partname = part.partname().clone();
            let source_blob = if retain_source
                && xml_minifier::audit::package::is_xml_part(partname.as_str(), part.content_type())
            {
                Some(part.payload_handle())
            } else {
                None
            };
            self.source_relationships_xml.remove(&partname);
            self.parts.insert(partname.clone(), part);
            if let Some(source_blob) = source_blob {
                self.source_xml_parts
                    .insert(partname, source_blob.payload().clone());
            }
        }
        checkpoint()?;
        Ok(())
    }

    /// Validate that a new part name would not replace or conflict with an existing part.
    ///
    /// # Errors
    /// Returns an error if `partname` duplicates or conflicts with an existing
    /// part name.
    pub fn validate_new_part_name(&self, partname: &PackURI) -> Result<()> {
        for existing in self.parts.keys() {
            if let Some(conflict) = existing.conflict_with(partname) {
                return Err(part_name_conflict_error(existing, partname, conflict));
            }
        }
        Ok(())
    }

    /// Remove a part by name, returning whether it existed.
    pub fn remove_part(&mut self, partname: &PackURI) -> bool {
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        self.content_types_source_state.reset();
        self.source_xml_parts.remove(partname);
        self.source_relationships_xml.remove(partname);
        self.part_bytes = None;
        self.parts.remove(partname).is_some()
    }

    /// Get an iterator over the name and metadata of every part.
    ///
    /// The item type carries a part's name, content type and relationships
    /// and has no route to its payload. A package opened from an owned source
    /// decodes a part's payload on first access, and that decode can fail;
    /// this iterator is infallible and has nowhere to report the refusal, so
    /// it must not be able to reach a payload (ADR 0030). Use
    /// [`Self::try_iter_parts`] when the iteration needs bytes.
    pub fn iter_parts(&self) -> impl Iterator<Item = PartMetadata<'_>> {
        self.parts.values().map(|b| PartMetadata::new(&**b))
    }

    /// Get a fallible iterator over all parts in the package.
    ///
    /// Each item forces that part's payload decode as it is yielded, so a
    /// caller reading payloads sees the same typed refusal
    /// [`Self::get_part`] would return. Iteration continues past an item
    /// that failed; the failing part has no payload and never acquires one.
    pub fn try_iter_parts(&self) -> impl Iterator<Item = Result<&dyn Part>> {
        self.parts.values().map(|b| {
            let part: &dyn Part = &**b;
            part.ensure_payload()?;
            Ok(part)
        })
    }

    /// Iterate every part without decoding any payload.
    ///
    /// Crate-internal passes that must observe a part's storage — publication
    /// planning, provenance capture, the signature audit's metadata scan —
    /// use this and decide for themselves whether a payload is needed.
    pub(crate) fn iter_parts_undecoded(&self) -> impl Iterator<Item = &dyn Part> {
        self.parts.values().map(|b| {
            let part: &dyn Part = &**b;
            part
        })
    }

    /// The names of the parts whose presence, content type, payload or
    /// relationships differ from `base`, sorted.
    ///
    /// Payloads are compared by allocation, never decoded: a package and the
    /// clone an edit derived from it share every untouched payload, so this
    /// finds what the edit replaced at a cost proportional to the part count.
    /// Two separately loaded copies of equal bytes compare as different.
    #[must_use]
    pub fn changed_parts(&self, base: &Self) -> Vec<PackURI> {
        let mut changed: Vec<PackURI> = self
            .parts
            .iter()
            .filter(|(name, part)| {
                base.parts.get(*name).is_none_or(|before| {
                    before.content_type() != part.content_type()
                        || !before
                            .payload_handle()
                            .payload()
                            .same_allocation(part.payload_handle().payload())
                        || !before.rels().same_entries(part.rels())
                })
            })
            .map(|(name, _)| name.clone())
            .collect();
        changed.extend(
            base.parts
                .keys()
                .filter(|name| !self.parts.contains_key(*name))
                .cloned(),
        );
        changed.sort_unstable_by(|left, right| left.as_str().cmp(right.as_str()));
        changed
    }

    /// Get the number of parts in the package.
    #[must_use]
    pub fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// Return the retained aggregate length of ordinary part payloads.
    ///
    /// A package assembled through a mutable escape hatch can temporarily
    /// report `None`; format owners refresh that value from their existing
    /// capture pass before publishing a package that needs replay admission.
    #[must_use]
    pub const fn admitted_part_bytes(&self) -> Option<u64> {
        self.part_bytes
    }

    /// Install an aggregate already proved by an ingress or materialization
    /// catalog pass. This is crate-private so callers cannot claim a metric
    /// for bytes they did not admit under the package's source policy.
    pub(crate) const fn set_admitted_part_bytes(&mut self, bytes: u64) {
        self.part_bytes = Some(bytes);
    }

    /// Parts inflated, and bytes inflated, from the retained source archive.
    ///
    /// `None` for a package whose payloads were all materialized at open —
    /// borrowed ingress, an explicitly scheduled eager open, or a package
    /// authored in memory. This reports what a lazy open has actually paid
    /// and exists for measurement and regression tests; it is not part of the
    /// supported surface.
    #[doc(hidden)]
    #[must_use]
    pub fn deferred_decode_counters(&self) -> Option<(u64, u64)> {
        self.parts.values().find_map(|part| {
            let handle = part.payload_handle();
            handle
                .payload()
                .deferred_source()
                .map(|source| source.counters())
        })
    }

    /// Get a reference to the package-level relationships.
    #[must_use]
    pub fn rels(&self) -> &Relationships {
        &self.rels
    }

    /// Get a mutable reference to the package-level relationships.
    pub fn rels_mut(&mut self) -> &mut Relationships {
        let was_signed = self.is_signed();
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.signature_graph_tracked = true;
        if !was_signed && !self.source_ingress {
            self.signature_policy_authorized = true;
        }
        &mut self.rels
    }

    /// Whether signature infrastructure is present anywhere in the package.
    ///
    /// This is a cheap capability check. It intentionally includes orphaned or
    /// partially formed signature parts and relationship targets, because a
    /// mutating writer must not treat an incomplete signature graph as an
    /// unsigned package. Opaque non-Part members whose names are rooted in the
    /// reserved signature directory are treated conservatively for the same
    /// reason. `unsign` remains the explicit authorization boundary for
    /// publishing a package after signature handling; opaque members are not
    /// exposed as raw payloads and are preserved as source archive entries.
    /// Use `signatures` when graph validation and cryptographic verification
    /// are required.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        self.rels.iter().any(is_signature_relationship_or_target)
            || self.parts.values().any(|part| {
                is_signature_infrastructure(&**part)
                    || part.rels().iter().any(is_signature_relationship_or_target)
            })
            || self
                .non_part_members
                .iter()
                .any(|member| is_signature_member_path(member.name()))
    }

    /// Verifies every OPC signature with the safe strict policy.
    ///
    /// # Errors
    /// Returns an error if the signature graph is ambiguous or spoofed, or if
    /// signature verification fails.
    #[cfg(feature = "sign")]
    pub fn signatures(&self) -> crate::sign::Result<Vec<crate::sign::Report>> {
        crate::sign::signatures(self, &litchi_sign::Policy::strict())
    }

    /// Verifies every OPC signature with an explicit trust-neutral policy.
    ///
    /// # Errors
    /// Returns an error if the signature graph is ambiguous or spoofed, or if
    /// signature verification fails.
    #[cfg(feature = "sign")]
    pub fn signatures_with(
        &self,
        policy: &litchi_sign::Policy,
    ) -> crate::sign::Result<Vec<crate::sign::Report>> {
        crate::sign::signatures(self, policy)
    }

    /// Adds a signature while retaining every existing valid signature.
    ///
    /// # Errors
    /// Returns an error if an existing signature is invalid or the new
    /// signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn sign(&mut self, signer: &litchi_sign::Signer) -> crate::sign::Result<PackURI> {
        let exact_source_authorized = self.exact_source_authorized;
        let signature_graph_tracked = self.signature_graph_tracked;
        let signature_policy_authorized = self.signature_policy_authorized;
        let signature_policy_required = self.signature_policy_required;
        let signature_api_authored = self.signature_api_authored;
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let result = crate::sign::sign(self, signer, &litchi_sign::Limits::standard());
        if result.is_ok() {
            self.signature_graph_tracked = true;
            self.signature_policy_authorized = true;
            self.signature_policy_required = true;
            self.signature_api_authored = true;
        } else {
            self.exact_source_authorized = exact_source_authorized;
            self.signature_graph_tracked = signature_graph_tracked;
            self.signature_policy_authorized = signature_policy_authorized;
            self.signature_policy_required = signature_policy_required;
            self.signature_api_authored = signature_api_authored;
        }
        result
    }

    /// Adds a signature with explicit authoring resource bounds.
    ///
    /// # Errors
    /// Returns an error if an existing signature is invalid, `limits` are
    /// exceeded, or the new signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn sign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> crate::sign::Result<PackURI> {
        let exact_source_authorized = self.exact_source_authorized;
        let signature_graph_tracked = self.signature_graph_tracked;
        let signature_policy_authorized = self.signature_policy_authorized;
        let signature_policy_required = self.signature_policy_required;
        let signature_api_authored = self.signature_api_authored;
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let result = crate::sign::sign(self, signer, limits);
        if result.is_ok() {
            self.signature_graph_tracked = true;
            self.signature_policy_authorized = true;
            self.signature_policy_required = true;
            self.signature_api_authored = true;
        } else {
            self.exact_source_authorized = exact_source_authorized;
            self.signature_graph_tracked = signature_graph_tracked;
            self.signature_policy_authorized = signature_policy_authorized;
            self.signature_policy_required = signature_policy_required;
            self.signature_api_authored = signature_api_authored;
        }
        result
    }

    /// Atomically replaces the validated signature graph with one signature.
    ///
    /// # Errors
    /// Returns an error if the signature graph is invalid or the replacement
    /// signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn resign(&mut self, signer: &litchi_sign::Signer) -> crate::sign::Result<PackURI> {
        let exact_source_authorized = self.exact_source_authorized;
        let signature_graph_tracked = self.signature_graph_tracked;
        let signature_policy_authorized = self.signature_policy_authorized;
        let signature_policy_required = self.signature_policy_required;
        let signature_api_authored = self.signature_api_authored;
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let result = crate::sign::resign(self, signer, &litchi_sign::Limits::standard());
        if result.is_ok() {
            self.signature_graph_tracked = true;
            self.signature_policy_authorized = true;
            self.signature_policy_required = true;
            self.signature_api_authored = true;
        } else {
            self.exact_source_authorized = exact_source_authorized;
            self.signature_graph_tracked = signature_graph_tracked;
            self.signature_policy_authorized = signature_policy_authorized;
            self.signature_policy_required = signature_policy_required;
            self.signature_api_authored = signature_api_authored;
        }
        result
    }

    /// Atomically replaces signatures with explicit authoring resource bounds.
    ///
    /// # Errors
    /// Returns an error if the signature graph is invalid, `limits` are exceeded,
    /// or the replacement signature cannot be created or staged into the package.
    #[cfg(feature = "sign")]
    pub fn resign_with(
        &mut self,
        signer: &litchi_sign::Signer,
        limits: &litchi_sign::Limits,
    ) -> crate::sign::Result<PackURI> {
        let exact_source_authorized = self.exact_source_authorized;
        let signature_graph_tracked = self.signature_graph_tracked;
        let signature_policy_authorized = self.signature_policy_authorized;
        let signature_policy_required = self.signature_policy_required;
        let signature_api_authored = self.signature_api_authored;
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let result = crate::sign::resign(self, signer, limits);
        if result.is_ok() {
            self.signature_graph_tracked = true;
            self.signature_policy_authorized = true;
            self.signature_policy_required = true;
            self.signature_api_authored = true;
        } else {
            self.exact_source_authorized = exact_source_authorized;
            self.signature_graph_tracked = signature_graph_tracked;
            self.signature_policy_authorized = signature_policy_authorized;
            self.signature_policy_required = signature_policy_required;
            self.signature_api_authored = signature_api_authored;
        }
        result
    }

    /// Removes all signature relationships and infrastructure parts.
    ///
    /// Deletion is infallible and idempotent, including for a malformed graph.
    pub fn unsign(&mut self) {
        self.revoke_exact_source();
        self.strip_signature_graph();
        self.signature_policy_authorized = true;
        self.signature_policy_required = false;
        self.signature_api_authored = false;
    }

    /// Relate the package to a part.
    ///
    /// Creates or reuses a relationship from the package to the specified part.
    ///
    /// # Arguments
    /// * `partname` - The target part's partname
    /// * `reltype` - The relationship type URI
    ///
    /// # Returns
    /// The relationship ID (rId)
    pub fn relate_to(&mut self, partname: &str, reltype: &str) -> String {
        let was_signed = self.is_signed();
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let r_id = self.rels.get_or_add(reltype, partname).r_id().to_string();
        self.signature_graph_tracked = true;
        if !was_signed && !self.source_ingress && is_signature_relationship(reltype) {
            self.signature_policy_authorized = true;
        }
        r_id
    }

    /// Add an external relationship (e.g., for hyperlinks).
    ///
    /// # Arguments
    /// * `target_url` - External URL target
    /// * `reltype` - Relationship type
    ///
    /// # Returns
    /// The relationship ID (e.g., "rId1")
    pub fn relate_to_external(&mut self, target_url: &str, reltype: &str) -> String {
        let was_signed = self.is_signed();
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let r_id = self.rels.get_or_add_ext_rel(reltype, target_url);
        self.signature_graph_tracked = true;
        if !was_signed && !self.source_ingress && is_signature_relationship(reltype) {
            self.signature_policy_authorized = true;
        }
        r_id
    }

    /// Get mutable access to package-level relationships.
    ///
    /// Useful for advanced relationship management.
    pub fn relationships_mut(&mut self) -> &mut Relationships {
        let was_signed = self.is_signed();
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        self.signature_graph_tracked = true;
        if !was_signed && !self.source_ingress {
            self.signature_policy_authorized = true;
        }
        &mut self.rels
    }

    /// Find the next available partname for a part template.
    ///
    /// Useful for creating new parts with sequential numbering (e.g., image1.png, image2.png).
    /// Uses efficient string operations to minimize allocations.
    ///
    /// # Arguments
    /// * `template` - A format string with a %d placeholder for the number
    ///
    /// # Example
    /// ```no_run
    /// # use litchi_opc::package::OpcPackage;
    /// # let mut pkg = OpcPackage::new();
    /// let next_image = pkg.next_partname("/word/media/image%d.png");
    /// ```
    ///
    /// # Errors
    /// Returns an error if the template has no `%d` placeholder, a candidate
    /// partname is invalid, or no free name exists within the bounded search.
    pub fn next_partname(&self, template: &str) -> Result<PackURI> {
        // Find the position of %d in the template for efficient replacement
        let percent_d_pos = template.find("%d").ok_or_else(|| {
            OpcError::InvalidPackUri("Template must contain %d placeholder".to_string())
        })?;

        let mut n = 1u32;
        let mut candidate_bytes = Vec::with_capacity(template.len() + 10); // Pre-allocate

        loop {
            // Clear and reuse the vector for each candidate
            candidate_bytes.clear();

            // Build candidate string more efficiently
            candidate_bytes.extend_from_slice(&template.as_bytes()[..percent_d_pos]);
            candidate_bytes.extend_from_slice(itoa::Buffer::new().format(n).as_bytes());
            candidate_bytes.extend_from_slice(&template.as_bytes()[percent_d_pos + 2..]);

            // Create PackURI from bytes to avoid intermediate string allocation
            let candidate_str = std::str::from_utf8(&candidate_bytes).map_err(|_err| {
                OpcError::InvalidPackUri("Invalid UTF-8 in partname".to_string())
            })?;

            let candidate_uri = PackURI::new(candidate_str).map_err(OpcError::InvalidPackUri)?;
            if !self.parts.contains_key(&candidate_uri) {
                return Ok(candidate_uri);
            }

            n += 1;
            if n > 10000 {
                // Safety limit to prevent infinite loops
                return Err(OpcError::InvalidPackUri(
                    "Too many parts, cannot find next partname".to_string(),
                ));
            }
        }
    }

    /// Check if a part exists in the package.
    #[must_use]
    pub fn contains_part(&self, partname: &PackURI) -> bool {
        self.parts.contains_key(partname)
    }

    /// Atomically save the package to a file.
    ///
    /// Writes and synchronizes a finalized sibling artifact before replacing
    /// the destination. A failure before replacement leaves it untouched.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::package::OpcPackage;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized or the file cannot
    /// be written; the destination is left untouched on failure before replacement.
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        crate::pkgwriter::PackageWriter::write(path, self)
    }

    /// Atomically save the package to a file at a caller-chosen
    /// [`Durability`](litchi_core::Durability).
    ///
    /// [`Self::save`] is this method at `Durability::Full`. Every level
    /// publishes the same bytes through the same sibling temporary file and
    /// rename, and leaves the destination untouched on a failure before the
    /// rename; see [`crate::atomic::replace_with_durability`] for what each
    /// level synchronizes. The level is not stored on the package, so
    /// choosing one changes no save preference and does not revoke the
    /// package's exact-source authorization.
    ///
    /// # Errors
    /// As [`Self::save`], except that only `Durability::Full` can return
    /// [`crate::OpcError::Committed`].
    pub fn save_with_durability<P: AsRef<Path>>(
        &self,
        path: P,
        durability: litchi_core::Durability,
    ) -> Result<()> {
        crate::pkgwriter::PackageWriter::write_with_durability(path, self, durability)
    }

    /// Save the package to a stream.
    ///
    /// Writes the complete OPC package including all parts, relationships,
    /// and content types directly to a writer stream. A failure can leave the
    /// caller-owned stream incomplete; the error reports accepted bytes.
    ///
    /// # Arguments
    /// * `writer` - Any sequential writer; seeking is not required
    ///
    /// # Example
    /// ```no_run
    /// use litchi_opc::package::OpcPackage;
    /// use std::fs::File;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// let file = File::create("output.docx")?;
    /// pkg.write_to(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized or written to the
    /// stream; the stream may be left incomplete.
    pub fn write_to<W: Write>(&self, writer: W) -> Result<()> {
        crate::pkgwriter::PackageWriter::write_to(writer, self)
    }

    pub(crate) fn strip_signature_graph(&mut self) {
        self.revoke_exact_source();
        self.content_types_source_state.reset();
        let infrastructure: HashSet<PackURI> = self
            .parts
            .values()
            .filter(|part| is_signature_infrastructure(&***part))
            .map(|part| part.partname().clone())
            .collect();

        self.rels.retain(|relationship| {
            !is_signature_relationship_or_target(relationship)
                && !targets_any(relationship, &infrastructure)
        });
        for part in self.parts.values_mut() {
            part.rels_mut().retain(|relationship| {
                !is_signature_relationship_or_target(relationship)
                    && !targets_any(relationship, &infrastructure)
            });
        }
        self.parts
            .retain(|_, part| !is_signature_infrastructure(&**part));
    }

    /// Shared handle to the exact owned source archive, while the
    /// exact-source authorization is intact.
    ///
    /// The returned handle borrows nothing: it is a further owner of the very
    /// allocation an owned ingress (`from_vec`, `from_vec_reusing_payloads`,
    /// `open`, `from_reader`) moved into this package, or a shared ingress
    /// (`from_shared_vec`, `from_shared_vec_reusing_payloads`)
    /// shared into it, so a caller may keep the archive alive after the
    /// package is dropped without copying it. The bytes are immutable through
    /// the handle and are exactly the bytes this package republishes
    /// verbatim.
    ///
    /// Returns `None` for a package that was not opened from owned bytes, and
    /// for one whose exact-source authorization has been revoked by an edit;
    /// a later revocation does not retract a handle already taken, so a caller
    /// that holds one must not treat it as evidence about the package's
    /// current state.
    #[must_use]
    pub fn exact_source_shared(&self) -> Option<Arc<Vec<u8>>> {
        self.exact_source_authorized
            .then(|| {
                self.source_archive
                    .as_ref()
                    .map(|source| Arc::clone(source.bytes()))
            })
            .flatten()
    }

    pub(crate) fn exact_source(&self) -> Option<&[u8]> {
        self.exact_source_authorized
            .then(|| {
                self.source_archive
                    .as_ref()
                    .map(|source| source.bytes().as_slice())
            })
            .flatten()
    }

    pub(crate) fn preservation_source(&self) -> Option<(&[u8], &PreservationProvenance)> {
        self.source_archive
            .as_ref()
            .map(|source| source.bytes().as_slice())
            .zip(self.preservation.as_deref())
    }

    pub(crate) fn mark_source_ingress_signature_policy(&mut self) {
        self.source_ingress = true;
        self.signature_graph_tracked = self.is_signed();
        self.signature_policy_required = self.signature_graph_tracked;
        self.signature_policy_authorized = false;
        self.signature_api_authored = false;
    }

    /// Whether a changed source still requires explicit signature disposition.
    ///
    /// This state can remain true after low-level removal of all visible
    /// signature parts or relationships. Format-owned publication must not
    /// authorize that change merely because [`Self::is_signed`] is now false.
    /// Use [`Self::unsign`] or the signing APIs to authorize the disposition.
    /// Exact unchanged source publication does not require a new disposition.
    #[must_use]
    pub fn requires_signature_edit_policy(&self) -> bool {
        !self.exact_source_authorized
            && self.signature_graph_tracked
            && ((self.signature_policy_required && !self.signature_policy_authorized)
                || (self.is_signed()
                    && (self.source_ingress || self.signature_api_authored)
                    && !self.signature_policy_authorized))
    }

    /// Validate signature disposition when publishing an edited candidate.
    ///
    /// A signed source with a retained archive may publish its unchanged clone
    /// or a clone explicitly handled through [`Self::unsign`] or the signing
    /// APIs. Replacing that candidate with a separately constructed package
    /// cannot discard the source's signature policy. Call `unsign` on the
    /// source before replacing its whole graph. Signed sources without a
    /// retained archive also require disposition before constructing a candidate.
    /// Changed candidates with opaque non-Part signature-directory entries are
    /// refused: `unsign` cannot prove removal of those retained archive members.
    ///
    /// This checks publication policy only, not cryptographic validity or trust.
    ///
    /// # Errors
    ///
    /// Returns [`OpcError::SignedSourceRequiresExplicitPolicy`] when either the
    /// candidate or its replacement of the source lacks explicit disposition.
    /// Returns [`OpcError::PreservationUnavailable`] for changed candidates
    /// retaining opaque signature-directory entries.
    pub fn validate_signature_edit_from(&self, source: &Self) -> Result<()> {
        if self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }
        if !self.exact_source_authorized
            && self
                .non_part_members
                .iter()
                .any(|member| is_signature_member_path(member.name()))
        {
            return Err(OpcError::PreservationUnavailable {
                reason: "opaque signature-directory entries cannot be removed by candidate signature disposition".to_owned(),
            });
        }
        if source.is_signed() || source.requires_signature_edit_policy() {
            let same_source = self
                .source_archive
                .as_ref()
                .zip(source.source_archive.as_ref())
                .is_some_and(|(candidate, original)| {
                    Arc::ptr_eq(candidate.bytes(), original.bytes())
                });
            if !same_source {
                return Err(OpcError::SignedSourceRequiresExplicitPolicy);
            }
        }
        Ok(())
    }

    pub(crate) fn requires_owned_source_preservation(&self) -> bool {
        self.source_archive.is_some() && !self.exact_source_authorized
    }

    fn revoke_exact_source(&mut self) {
        self.exact_source_authorized = false;
        if self.source_ingress || self.signature_api_authored {
            self.signature_policy_authorized = false;
            if self.is_signed() {
                self.signature_policy_required = true;
            }
        }
    }

    fn from_owned_bytes_with_limits(data: Vec<u8>, limits: SourceLimits) -> Result<Self> {
        Self::from_shared_bytes_with_limits(Arc::new(data), limits)
    }

    fn from_shared_bytes_with_limits(data: Arc<Vec<u8>>, limits: SourceLimits) -> Result<Self> {
        let mut package = {
            let phys_reader = PhysPkgReader::new(data.as_slice(), limits)?;
            let pkg_reader = PackageReader::from_phys_reader_deferred(&phys_reader, &data)?;
            Self::unmarshal(pkg_reader)?
        };
        package.authorize_shared_owned_source(data, limits);
        Ok(package)
    }

    pub(crate) fn from_bytes_with_open_session(
        data: &[u8],
        limits: SourceLimits,
        session: &OpenSession,
    ) -> Result<Self> {
        session.check()?;
        let phys_reader = PhysPkgReader::new(data, limits)?;
        session.charge_input(data.len() as u64)?;
        let pkg_reader = PackageReader::from_phys_reader_with_session(&phys_reader, session)?;
        Self::unmarshal(pkg_reader)
    }

    pub(crate) fn from_vec_with_open_session(
        data: Vec<u8>,
        limits: SourceLimits,
        session: &OpenSession,
    ) -> Result<Self> {
        session.check()?;
        let input_bytes = data.len() as u64;
        let mut package = {
            let phys_reader = PhysPkgReader::new(&data, limits)?;
            session.charge_input(input_bytes)?;
            let pkg_reader = PackageReader::from_phys_reader_with_session(&phys_reader, session)?;
            Self::unmarshal(pkg_reader)?
        };
        package.authorize_owned_source(data, limits);
        Ok(package)
    }

    fn authorize_owned_source(&mut self, source: Vec<u8>, limits: SourceLimits) {
        self.authorize_shared_owned_source(Arc::new(source), limits);
    }

    /// Retain an owned source archive that deferred payloads already share.
    fn authorize_shared_owned_source(&mut self, source: Arc<Vec<u8>>, limits: SourceLimits) {
        let source_content_types_xml = self.source_content_types_xml.clone();
        let source_content_types = self.source_content_types.clone();
        let _ = self.attach_owned_source_preservation(
            source,
            source_content_types_xml,
            source_content_types,
            limits,
            true,
        );
    }

    /// Attach source archive/provenance to a package materialized from an
    /// already admitted source-backed catalog.
    ///
    /// The source bytes and parsed content-types binding are transferred by
    /// shared ownership.  The OPC catalog is not reparsed; a bounded ZIP
    /// member scan builds the physical publication provenance.  `exact` is
    /// true only for a graph that is still the physical source; changed
    /// candidates retain the same archive and provenance for targeted
    /// publication while forcing the writer through its changed-package path.
    pub(crate) fn attach_owned_source_preservation(
        &mut self,
        source: Arc<Vec<u8>>,
        source_content_types_xml: Option<Arc<Vec<u8>>>,
        source_content_types: Option<Arc<ContentTypeMap>>,
        limits: SourceLimits,
        exact: bool,
    ) -> Result<()> {
        if let (Some(xml), Some(binding)) = (source_content_types_xml, source_content_types) {
            self.source_content_types_xml = Some(xml);
            self.source_content_types = Some(binding);
            self.content_types_source_state.reset();
        }
        let preservation = PreservationProvenance::from_package(source.as_slice(), self);
        if !exact && preservation.is_none() {
            return Err(OpcError::PreservationUnavailable {
                reason: "owned source preservation provenance could not be established".to_owned(),
            });
        }
        if let Some(preservation) = preservation.as_ref() {
            self.bind_relationship_captures(preservation);
        }
        self.preservation = preservation.map(Arc::new);
        self.source_archive = Some(retained_archive::RetainedArchive::new(source));
        self.source_limits = limits;
        self.transfer_index = Some(Arc::new(OnceLock::new()));
        self.exact_source_authorized = exact;
        Ok(())
    }

    /// Hand every relationship collection the canonical serialization the
    /// provenance captured from it.
    ///
    /// Publication planning later compares the two handles by pointer
    /// identity. A part whose relationships have not changed since this point
    /// therefore proves itself pristine without reserializing, so its source
    /// member is copied instead of rebuilt. Only parts whose source archive
    /// actually carries a relationships member are bound, because a part whose
    /// member must be created still has to be serialized and audited.
    fn bind_relationship_captures(&mut self, preservation: &PreservationProvenance) {
        for (partname, part) in &mut self.parts {
            if let Some(source_part) = preservation.parts.get(partname)
                && source_part.relationships_member_present
            {
                part.rels_mut()
                    .set_source_capture(Arc::clone(&source_part.relationships_xml));
            }
        }
        self.rels
            .set_source_capture(Arc::clone(&preservation.package_relationships_xml));
    }
}

/// The retained owned source archive and the digest memo bound to it.
///
/// The fields are private to this module and [`RetainedArchive::new`] is the only
/// constructor, so an archive always enters a package with a fresh, empty
/// memo, and a digest can never be attached to other bytes: struct-update
/// syntax or field assignment elsewhere in the crate cannot pair one archive's
/// digest with another archive. The memo is filled at most once, from exactly
/// the bytes it describes; clones share both the archive and the memo, so an
/// ingress and every clone of it hash the archive at most once between them.
mod retained_archive {
    use std::sync::{Arc, OnceLock};

    use sha2::{Digest as _, Sha256};

    #[derive(Clone)]
    pub(super) struct RetainedArchive {
        /// The retained archive. It is never mutated: no path reaches it
        /// through `Arc::get_mut` or `Arc::make_mut`, and a further owner of
        /// the allocation cannot mutate it in place while this one exists.
        bytes: Arc<Vec<u8>>,
        /// SHA-256 of `bytes`, computed on first request.
        sha256: Arc<OnceLock<[u8; 32]>>,
    }

    impl RetainedArchive {
        pub(super) fn new(bytes: Arc<Vec<u8>>) -> Self {
            Self {
                bytes,
                sha256: Arc::new(OnceLock::new()),
            }
        }

        pub(super) const fn bytes(&self) -> &Arc<Vec<u8>> {
            &self.bytes
        }

        pub(super) fn sha256(&self) -> [u8; 32] {
            let digest = *self
                .sha256
                .get_or_init(|| Sha256::digest(self.bytes.as_slice()).into());
            // Test and debug builds re-derive every read, so the suite proves
            // the memo describes its own bytes rather than only that it
            // compiles.
            debug_assert_eq!(
                digest,
                <[u8; 32]>::from(Sha256::digest(self.bytes.as_slice())),
                "an owned source's digest memo answered for different bytes"
            );
            digest
        }

        /// Whether the digest has been computed, for the memo tests.
        #[cfg(test)]
        pub(super) fn sha256_is_memoized(&self) -> bool {
            self.sha256.get().is_some()
        }

        /// Whether two values share one memo, for the memo tests.
        #[cfg(test)]
        pub(super) fn shares_memo_with(&self, other: &Self) -> bool {
            Arc::ptr_eq(&self.sha256, &other.sha256)
        }
    }
}

mod compressed_transfer;
pub use compressed_transfer::CompressedPartTransfer;
use compressed_transfer::TransferIndexCell;

#[cfg(test)]
mod payload_reuse_tests;

impl Default for OpcPackage {
    fn default() -> Self {
        Self::new()
    }
}

impl PreservationProvenance {
    fn from_package(source: &[u8], package: &OpcPackage) -> Option<Self> {
        let content_types_xml = package.source_content_types_xml.as_ref()?.clone();
        let archive = soapberry_zip::ZipArchive::from_slice(source).ok()?;
        let entry_count = usize::try_from(archive.entries_hint()).ok()?;

        let mut actual_entry_count = 0usize;
        for entry in archive.entries() {
            entry.ok()?;
            actual_entry_count = actual_entry_count.checked_add(1)?;
        }
        if actual_entry_count != entry_count {
            return None;
        }

        let mut parts = HashMap::new();
        parts.try_reserve(package.part_count()).ok()?;
        let mut part_members = HashMap::new();
        part_members.try_reserve(package.part_count()).ok()?;
        let mut relationship_members = HashMap::new();
        relationship_members
            .try_reserve(package.part_count())
            .ok()?;
        for part in package.iter_parts_undecoded() {
            let partname = part.partname().clone();
            let member_name = try_owned_string(part.partname().membername())?;
            if part_members.insert(member_name, partname.clone()).is_some() {
                return None;
            }
            let relationships_uri = part.partname().rels_uri().ok()?;
            if relationship_members
                .insert(
                    try_owned_string(relationships_uri.membername())?,
                    partname.clone(),
                )
                .is_some()
            {
                return None;
            }
            parts.insert(
                partname,
                SourcePart {
                    content_type: try_owned_string(part.content_type())?,
                    blob: part.payload_handle().payload().clone(),
                    relationships_xml: Arc::new(PreservedRelationshipsXml::from_package(
                        package,
                        part.partname(),
                        part.rels(),
                    )?),
                    member_present: false,
                    relationships_member_present: false,
                },
            );
        }

        let package_uri = PackURI::new(PACKAGE_URI).ok()?;
        let package_relationships_name = package_uri.rels_uri().ok()?.membername().to_owned();
        let mut content_types_present = false;
        let mut package_relationships_present = false;
        let mut seen_names = HashSet::new();
        seen_names.try_reserve(entry_count).ok()?;
        let mut members = Vec::new();
        members.try_reserve_exact(entry_count).ok()?;
        for entry in archive.entries() {
            let entry = entry.ok()?;
            let raw_name = entry.file_path();
            let raw_name = raw_name.as_ref();
            if !seen_names.insert(try_owned_bytes(raw_name)?) {
                return None;
            }
            let Ok(name) = std::str::from_utf8(raw_name) else {
                members.push(SourceMember {
                    kind: SourceMemberKind::Unknown,
                });
                continue;
            };

            let kind = if name.eq_ignore_ascii_case("[Content_Types].xml") {
                if content_types_present {
                    return None;
                }
                content_types_present = true;
                SourceMemberKind::ContentTypes
            } else if name == package_relationships_name {
                if package_relationships_present {
                    return None;
                }
                package_relationships_present = true;
                SourceMemberKind::PackageRelationships
            } else if let Some(partname) = part_members.get(name) {
                let part = parts.get_mut(partname)?;
                if part.member_present {
                    return None;
                }
                part.member_present = true;
                SourceMemberKind::Part(partname.clone())
            } else if let Some(partname) = relationship_members.get(name) {
                let part = parts.get_mut(partname)?;
                if part.relationships_member_present {
                    return None;
                }
                part.relationships_member_present = true;
                SourceMemberKind::PartRelationships(partname.clone())
            } else {
                SourceMemberKind::Unknown
            };
            members.push(SourceMember { kind });
        }

        if members.len() != entry_count
            || !content_types_present
            || parts.values().any(|part| !part.member_present)
        {
            return None;
        }

        Some(Self {
            members,
            parts,
            content_types_xml,
            package_relationships_xml: Arc::new(PreservedRelationshipsXml::from_package(
                package,
                &package_uri,
                package.rels(),
            )?),
        })
    }
}

/// Audit caller-supplied XML exactly as the package writer would publish it,
/// returning the proof a part carries so the writer need not audit it again.
fn verified_source_payload(
    name: &PackURI,
    bytes: Arc<Vec<u8>>,
) -> Result<xml_minifier::audit::VerifiedSource> {
    verified_source_payload_with_limits(name, bytes, xml_minifier::audit::Limits::default())
}

fn verified_source_payload_with_limits(
    name: &PackURI,
    bytes: Arc<Vec<u8>>,
    limits: xml_minifier::audit::Limits,
) -> Result<xml_minifier::audit::VerifiedSource> {
    xml_minifier::audit::VerifiedSource::verify(bytes, limits).map_err(|source| {
        OpcError::XmlPublication {
            part: name.to_string(),
            source,
        }
    })
}

fn verified_source_payload_with_source_limits(
    name: &PackURI,
    bytes: Arc<Vec<u8>>,
    limits: SourceLimits,
) -> Result<xml_minifier::audit::VerifiedSource> {
    let audit_limits = xml_minifier::audit::Limits::default()
        .narrow(
            xml_minifier::audit::Resource::Bytes,
            as_usize(limits.max_part_bytes()),
        )
        .narrow(xml_minifier::audit::Resource::Depth, limits.max_depth())
        .narrow(xml_minifier::audit::Resource::Events, limits.max_events())
        .narrow(
            xml_minifier::audit::Resource::TokenBytes,
            as_usize(limits.max_part_bytes()),
        )
        .narrow(
            xml_minifier::audit::Resource::TextBytes,
            as_usize(limits.max_part_bytes()),
        );
    verified_source_payload_with_limits(name, bytes, audit_limits)
}

const fn as_usize(value: u64) -> usize {
    if value > usize::MAX as u64 {
        usize::MAX
    } else {
        value as usize
    }
}

/// Audit a `[Content_Types].xml` or `.rels` token the package is about to
/// retain as its source, under limits that honour the read policy.
///
/// The limits are the writer's defaults, except that the input, token and
/// character-data budgets grow to cover the token, up to `admitted` (the read
/// policy's limit for that member) and the audit's hard ceilings. A caller
/// that explicitly raised the manifest or relationships limit can then install
/// the source-derived token it was allowed to read. Every structural check
/// still runs, and a token larger than `admitted` gets the default budgets
/// and is refused.
fn verified_metadata_token(
    name: &PackURI,
    bytes: Arc<Vec<u8>>,
    admitted: usize,
) -> Result<xml_minifier::audit::VerifiedSource> {
    use xml_minifier::audit::Limits;
    let defaults = Limits::default();
    let covered = if bytes.len() <= admitted {
        bytes.len()
    } else {
        0
    };
    let grown = Limits::builder()
        .bytes(defaults.max_bytes().max(covered.min(Limits::BYTE_CEILING)))
        .and_then(|builder| {
            builder.token_bytes(
                defaults
                    .max_token_bytes()
                    .max(covered.min(Limits::TOKEN_BYTE_CEILING)),
            )
        })
        .and_then(|builder| {
            builder.text_bytes(
                defaults
                    .max_text_bytes()
                    .max(covered.min(Limits::TEXT_BYTE_CEILING)),
            )
        })
        .map_or(defaults, |builder| builder.build());
    verified_source_payload_with_limits(name, bytes, grown)
}

fn try_owned_string(value: &str) -> Option<String> {
    let mut owned = String::new();
    owned.try_reserve_exact(value.len()).ok()?;
    owned.push_str(value);
    Some(owned)
}

fn try_owned_bytes(value: &[u8]) -> Option<Vec<u8>> {
    let mut owned = Vec::new();
    owned.try_reserve_exact(value.len()).ok()?;
    owned.extend_from_slice(value);
    Some(owned)
}

fn targets_any(relationship: &crate::Relationship, infrastructure: &HashSet<PackURI>) -> bool {
    if relationship.is_external() {
        return false;
    }
    match relationship.target_partname() {
        Ok(target) => infrastructure
            .iter()
            .any(|part| part.as_str().eq_ignore_ascii_case(target.as_str())),
        Err(_) => is_signature_path(relationship.target_path()),
    }
}

fn is_signature_relationship(kind: &str) -> bool {
    [
        relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate",
    ]
    .iter()
    .any(|candidate| kind.eq_ignore_ascii_case(candidate))
}

fn is_signature_relationship_or_target(relationship: &crate::Relationship) -> bool {
    if is_signature_relationship(relationship.reltype()) {
        return true;
    }
    if relationship.is_external() {
        return false;
    }
    let target_path = relationship.target_path();
    if is_signature_path(target_path) {
        return true;
    }
    if !target_may_be_signature_path(target_path) {
        return false;
    }
    relationship
        .target_partname()
        .map_or(true, |target| is_signature_path(target.as_str()))
}

fn target_may_be_signature_path(path: &str) -> bool {
    path.split('/')
        .any(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
}

fn is_signature_path(path: &str) -> bool {
    const DIRECTORY: &[u8] = b"/_xmlsignatures/";
    path.as_bytes()
        .get(..DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(DIRECTORY))
}

/// Return whether an opaque ZIP member name is rooted in the reserved OPC
/// signature directory.
///
/// Part names are absolute (`/_xmlsignatures/...`) while ZIP member names are
/// relative (`_xmlsignatures/...`). Non-Part names can also contain malformed
/// separators or harmless leading dot segments, so compare their first
/// meaningful path component using both separator styles. Only that root
/// component is accepted; a nested `_xmlsignatures` directory is unrelated to
/// the package signature directory.
fn is_signature_member_path(path: &str) -> bool {
    path.split(['/', '\\'])
        .find(|segment| !segment.is_empty() && *segment != ".")
        .is_some_and(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
}

fn is_signature_infrastructure(part: &dyn Part) -> bool {
    use crate::constants::content_type;

    is_signature_path(part.partname().as_str())
        || [
            content_type::OPC_DIGITAL_SIGNATURE_ORIGIN,
            content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE,
            content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE,
        ]
        .iter()
        .any(|candidate| part.content_type().eq_ignore_ascii_case(candidate))
}

fn part_name_conflict_error(
    existing: &PackURI,
    candidate: &PackURI,
    conflict: PartNameConflict,
) -> OpcError {
    match conflict {
        PartNameConflict::Duplicate => OpcError::DuplicatePartName(candidate.to_string()),
        PartNameConflict::Equivalent => OpcError::EquivalentPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
        PartNameConflict::Derived => OpcError::DerivedPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;
    use crate::part::BlobPart;
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::fs;
    use std::io::Cursor;
    use std::sync::Arc;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use tempfile::NamedTempFile;

    fn create_minimal_docx() -> Vec<u8> {
        create_minimal_docx_with_stored_document(false)
    }

    fn create_mixed_minimal_docx() -> Vec<u8> {
        create_minimal_docx_with_stored_document(true)
    }

    fn create_minimal_docx_with_stored_document(stored_document: bool) -> Vec<u8> {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());

        // Add [Content_Types].xml
        writer
            .write_deflated(
                "[Content_Types].xml",
                br#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#,
            )
            .unwrap();

        // Add _rels/.rels
        writer
            .write_deflated(
                "_rels/.rels",
                br#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#,
            )
            .unwrap();

        // Add word/document.xml
        let document = br#"<?xml version="1.0"?>
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <body><p><t>Test</t></p></body>
</document>"#;
        if stored_document {
            writer.write_stored("word/document.xml", document).unwrap();
        } else {
            writer
                .write_deflated("word/document.xml", document)
                .unwrap();
        }

        writer.finish_to_bytes().unwrap()
    }

    fn create_source_without_package_relationships() -> Vec<u8> {
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer.write_stored("document.xml", b"<document/>").unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn create_two_part_source() -> Vec<u8> {
        let mut package =
            OpcPackage::from_bytes(&create_minimal_docx(), SourceLimits::default()).unwrap();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/second.xml").unwrap(),
            crate::constants::content_type::XML.to_owned(),
            b"<second/>".to_vec(),
        )));
        crate::PackageWriter::to_bytes(&package).unwrap()
    }

    #[derive(Clone)]
    struct DecodeRefusingPart {
        partname: PackURI,
        payload: Arc<Vec<u8>>,
        rels: Relationships,
    }

    impl Part for DecodeRefusingPart {
        fn blob(&self) -> &[u8] {
            self.payload.as_slice()
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            Arc::clone(&self.payload)
        }

        fn content_type(&self) -> &str {
            crate::constants::content_type::XML
        }

        fn partname(&self) -> &PackURI {
            &self.partname
        }

        fn ensure_payload(&self) -> Result<()> {
            Err(OpcError::ZipError("test decode refusal".to_owned()))
        }

        fn rels(&self) -> &Relationships {
            &self.rels
        }

        fn rels_mut(&mut self) -> &mut Relationships {
            &mut self.rels
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.payload = Arc::new(blob);
        }
    }

    #[derive(Clone)]
    struct StatefulDecodePart {
        partname: PackURI,
        payload: Arc<Vec<u8>>,
        rels: Relationships,
        ensure_calls: Arc<AtomicUsize>,
        fail_on_call: usize,
    }

    impl Part for StatefulDecodePart {
        fn blob(&self) -> &[u8] {
            self.payload.as_slice()
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            Arc::clone(&self.payload)
        }

        fn content_type(&self) -> &str {
            crate::constants::content_type::XML
        }

        fn partname(&self) -> &PackURI {
            &self.partname
        }

        fn ensure_payload(&self) -> Result<()> {
            let call = self.ensure_calls.fetch_add(1, Ordering::SeqCst) + 1;
            if call == self.fail_on_call {
                Err(OpcError::ZipError(
                    "test stateful decode refusal".to_owned(),
                ))
            } else {
                Ok(())
            }
        }

        fn rels(&self) -> &Relationships {
            &self.rels
        }

        fn rels_mut(&mut self) -> &mut Relationships {
            &mut self.rels
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.payload = Arc::new(blob);
        }
    }

    fn large_content_types_archive() -> (Vec<u8>, Vec<u8>) {
        let mut content_types = Vec::new();
        content_types.extend_from_slice(
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><!--"#,
        );
        content_types.resize(
            content_types.len() + SourceLimits::default().max_content_types_bytes() + 1,
            b'x',
        );
        content_types.extend_from_slice(
            br#"--><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/document.xml" ContentType="application/xml"/></Types>"#,
        );

        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored("[Content_Types].xml", &content_types)
            .unwrap();
        writer
            .write_stored(
                "_rels/.rels",
                br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="document.xml"/></Relationships>"#,
            )
            .unwrap();
        writer.write_stored("document.xml", b"<document/>").unwrap();
        (writer.finish_to_bytes().unwrap(), content_types)
    }

    fn create_source_with_explicit_relationship_overrides(empty: bool) -> (Vec<u8>, Vec<u8>) {
        let content_types = br#"<?xml version='1.0'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <!-- explicit relationship declarations are intentionally lexical -->
  <Default Extension='bin' ContentType='application/octet-stream'/>
  <Default Extension='xml' ContentType='application/xml'/>
  <Override PartName='/_RELS/.RELS' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Override PartName='/WORD/_RELS/DOCUMENT.XML.RELS' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
</Types>
"#;
        let package_relationships = br#"<?xml version='1.0'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
</Relationships>
"#;
        let document_relationships: &[u8] = if empty {
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#
        } else {
            br#"<?xml version='1.0'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='urn:test:media' Target='../media.bin'/>
</Relationships>
"#
        };
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        for (name, bytes) in [
            ("[Content_Types].xml", content_types.as_slice()),
            ("_rels/.rels", package_relationships.as_slice()),
            ("word/document.xml", br#"<document/>"#.as_slice()),
            ("word/_rels/document.xml.rels", document_relationships),
            ("media.bin", b"original media".as_slice()),
        ] {
            writer.write_stored(name, bytes).unwrap();
        }
        (writer.finish_to_bytes().unwrap(), content_types.to_vec())
    }

    fn create_source_with_unused_relationship_override() -> (Vec<u8>, Vec<u8>) {
        let content_types = br#"<?xml version='1.0'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
  <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
  <Default Extension='xml' ContentType='application/xml'/>
  <Override PartName='/word/_rels/document.xml.rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
</Types>
"#;
        let package_relationships = br#"<?xml version='1.0'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
  <Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
</Relationships>
"#;
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        for (name, bytes) in [
            ("[Content_Types].xml", content_types.as_slice()),
            ("_rels/.rels", package_relationships.as_slice()),
            ("word/document.xml", br#"<document/>"#.as_slice()),
        ] {
            writer.write_stored(name, bytes).unwrap();
        }
        (writer.finish_to_bytes().unwrap(), content_types.to_vec())
    }

    fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Vec<u8> {
        let comment_len = u16::try_from(comment.len()).expect("ZIP comment fits in EOCD");
        let eocd = archive.len().checked_sub(22).expect("archive has an EOCD");
        assert_eq!(&archive[eocd..eocd + 4], b"PK\x05\x06");
        archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
        archive.extend_from_slice(comment);
        archive
    }

    #[test]
    fn test_open_package() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor, SourceLimits::default()).unwrap();

        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn source_without_package_relationships_charges_generated_root_metadata() {
        let source = create_source_without_package_relationships();
        let source_archive = soapberry_zip::office::ArchiveReader::new(
            &source,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert!(
            !source_archive
                .file_names()
                .any(|name| name == "_rels/.rels")
        );

        let package = OpcPackage::from_vec(source.clone(), SourceLimits::default()).unwrap();
        let metrics = package
            .retained_archive_metrics(SourceLimits::default())
            .unwrap()
            .expect("owned source metrics");
        let source_zip = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
        let source_entry_metadata: u64 = source_zip
            .entries()
            .map(|entry| entry.unwrap().metadata_size_hint())
            .sum();
        let expected_metadata = archive_tail_metadata_bytes(&source_zip).unwrap()
            + source_entry_metadata
            + "_rels/.rels".len() as u64;
        assert_eq!(metrics.metadata_bytes(), expected_metadata);

        let exact = SourceLimits::builder()
            .max_archive_metadata_bytes(metrics.metadata_bytes())
            .unwrap()
            .build()
            .unwrap();
        assert!(package.retained_archive_metrics(exact).is_ok());

        let under = SourceLimits::builder()
            .max_archive_metadata_bytes(metrics.metadata_bytes() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.retained_archive_metrics(under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::ArchiveMetadataBytes,
                ..
            })
        ));

        // A graph-only reopen exercises the ordinary writer path that must
        // synthesize the missing root member; the owned package above is kept
        // intact so the metrics cover its retained source archive.
        let authored = OpcPackage::from_bytes(&source, SourceLimits::default()).unwrap();
        let output = crate::PackageWriter::to_bytes(&authored).unwrap();
        let output_archive = soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert!(
            output_archive
                .file_names()
                .any(|name| name == "_rels/.rels")
        );
    }

    #[test]
    fn relationship_source_at_raw_limit_does_not_use_expanding_canonical_cache() {
        let target = format!("https://example.test/{}", ">".repeat(2048));
        let relationships = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="{target}" TargetMode="External"/></Relationships>"#
        );
        let mut writer =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", relationships.as_bytes())
            .unwrap();
        writer
            .write_stored("word/document.xml", b"<document/>")
            .unwrap();
        let source = writer.finish_to_bytes().unwrap();
        let limits = SourceLimits::builder()
            .max_relationship_xml_bytes(relationships.len())
            .unwrap()
            .max_total_relationship_xml_bytes(relationships.len())
            .unwrap()
            .build()
            .unwrap();

        let package = OpcPackage::from_bytes(&source, limits).unwrap();
        let owner = PackURI::new(PACKAGE_URI).unwrap();
        let token = package
            .source_relationships(&owner, SourceLimits::default())
            .unwrap();
        assert_eq!(token.bytes(), relationships.as_bytes());
    }

    #[test]
    fn retained_content_type_fast_comparison_tracks_coverage_changes() {
        for empty in [false, true] {
            let (bytes, _) = create_source_with_explicit_relationship_overrides(empty);
            let base = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
            let owner = PackURI::new("/word/document.xml").unwrap();
            let mut payload = base.clone();
            payload
                .get_part_mut(&owner)
                .unwrap()
                .set_blob(b"<changed/>".to_vec());
            let mut content_type = base.clone();
            content_type
                .get_part_mut(&owner)
                .unwrap()
                .set_content_type("application/example+xml".to_owned())
                .unwrap();
            let mut removed = base.clone();
            removed.remove_part(&owner);
            let mut presence = base.clone();
            presence
                .get_part_mut(&owner)
                .unwrap()
                .rels_mut()
                .retain(|_| false);
            presence.source_relationships_xml.remove(&owner);
            for changed in [&base, &payload, &content_type, &removed, &presence] {
                let expected = same_optional_manifest_bytes(
                    base.source_content_types_source()
                        .unwrap()
                        .map(|(bytes, _)| bytes),
                    changed
                        .source_content_types_source()
                        .unwrap()
                        .map(|(bytes, _)| bytes),
                );
                assert_eq!(base.same_retained_content_types(changed).unwrap(), expected);
                assert_eq!(
                    changed.same_retained_content_types(&base).unwrap(),
                    expected
                );
            }
        }
    }

    #[test]
    fn content_types_source_state_is_clone_local_and_tracks_mutation_routes() {
        let source = create_minimal_docx();
        let document = PackURI::new("/word/document.xml").unwrap();
        let extra = PackURI::new("/word/extra.xml").unwrap();
        let base = OpcPackage::from_bytes(&source, SourceLimits::default()).unwrap();
        assert_eq!(base.content_types_source_state.load(), None);
        assert!(base.retained_content_types_xml().unwrap().is_some());
        assert_eq!(base.content_types_source_state.load(), Some(true));

        // A clone owns an independent memo. Mutating its payload resets only
        // the clone, while the source package keeps its proven binding.
        let mut payload = base.clone();
        payload
            .get_part_mut(&document)
            .unwrap()
            .set_blob(b"<changed/>".to_vec());
        assert_eq!(payload.content_types_source_state.load(), None);
        assert_eq!(base.content_types_source_state.load(), Some(true));
        assert!(payload.retained_content_types_xml().unwrap().is_some());
        assert_eq!(payload.content_types_source_state.load(), Some(true));

        // Adding a new default-typed part with no relationship member resets
        // the memo; recomputation proves the admitted source remains valid.
        let mut added = base.clone();
        added.add_part(Box::new(BlobPart::new(
            extra,
            crate::constants::content_type::XML.to_owned(),
            b"<extra/>".to_vec(),
        )));
        assert_eq!(added.content_types_source_state.load(), None);
        assert!(added.retained_content_types_xml().unwrap().is_some());
        assert_eq!(added.content_types_source_state.load(), Some(true));

        // Replacing an existing URI with a different content type invalidates
        // the memo and cannot retain the old lexical manifest.
        let mut replaced = base.clone();
        replaced.add_part(Box::new(BlobPart::new(
            document.clone(),
            "application/example+xml".to_owned(),
            b"<replacement/>".to_vec(),
        )));
        assert_eq!(replaced.content_types_source_state.load(), Some(false));
        assert!(replaced.retained_content_types_xml().unwrap().is_none());
        assert_eq!(replaced.content_types_source_state.load(), Some(false));

        // A rejected duplicate must not memoize an invalid result for the
        // unchanged package.
        let mut rejected = base.clone();
        assert!(
            rejected
                .try_add_part(Box::new(BlobPart::new(
                    document.clone(),
                    "application/example+xml".to_owned(),
                    b"<rejected/>".to_vec(),
                )))
                .is_err()
        );
        assert!(rejected.retained_content_types_xml().unwrap().is_some());

        // A part relationship invalidates the memo. The source's `rels`
        // default covers the newly present member, so recomputation still
        // proves that the retained manifest remains authoritative.
        let mut part_relationship = base.clone();
        part_relationship
            .get_part_mut(&document)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:test:relationship".to_owned(),
                "../extra.xml".to_owned(),
                "rId1".to_owned(),
                crate::TargetMode::Internal,
            )
            .unwrap();
        assert_eq!(part_relationship.content_types_source_state.load(), None);
        assert!(
            part_relationship
                .retained_content_types_xml()
                .unwrap()
                .is_some()
        );
        assert_eq!(
            part_relationship.content_types_source_state.load(),
            Some(true)
        );

        // Package-level relationship edits are covered by the same reset,
        // but remain content-type compatible because the member is already
        // declared by the source's `rels` default.
        let mut root_relationship = base.clone();
        root_relationship
            .rels_mut()
            .try_add_relationship(
                "urn:test:external".to_owned(),
                "https://example.invalid/".to_owned(),
                "rId2".to_owned(),
                crate::TargetMode::External,
            )
            .unwrap();
        assert_eq!(root_relationship.content_types_source_state.load(), None);
        assert!(
            root_relationship
                .retained_content_types_xml()
                .unwrap()
                .is_some()
        );
        assert_eq!(
            root_relationship.content_types_source_state.load(),
            Some(true)
        );

        let mut removed = base.clone();
        assert!(removed.remove_part(&document));
        assert_eq!(removed.content_types_source_state.load(), None);
        assert!(removed.retained_content_types_xml().unwrap().is_none());

        // Unsigning runs the signature-graph removal path even when this
        // source has no signature infrastructure. Its reset is therefore
        // independently covered without requiring a signing key fixture.
        let mut unsigned = base.clone();
        unsigned.unsign();
        assert_eq!(unsigned.content_types_source_state.load(), None);
        assert!(unsigned.retained_content_types_xml().unwrap().is_some());
    }

    #[test]
    fn compact_manifest_views_match_tokens_and_bound_authored_output() {
        let mut package =
            OpcPackage::from_bytes(&create_minimal_docx(), SourceLimits::default()).unwrap();
        let limits = package.read_limits();
        let root = PackURI::new("/").unwrap();
        assert_eq!(
            package
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            package.source_content_types(limits).unwrap().bytes(),
        );
        let (bytes, present) = package.retained_relationships_xml(&root).unwrap().unwrap();
        let token = package.source_relationships(&root, limits).unwrap();
        assert_eq!(bytes.as_slice(), token.bytes());
        assert_eq!(present, token.member_present());
        let original = package.clone();
        package
            .get_part_mut(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .set_content_type("application/example+xml".to_owned())
            .unwrap();
        package
            .rels_mut()
            .try_add_relationship(
                "urn:example:external".to_owned(),
                "https://example.invalid/".to_owned(),
                "external".to_owned(),
                crate::TargetMode::External,
            )
            .unwrap();
        assert_eq!(
            package.content_types_source_bytes(),
            original.content_types_source_bytes()
        );
        assert!(package.retained_content_types_xml().unwrap().is_none());
        assert!(package.retained_relationships_xml(&root).unwrap().is_none());
        assert_eq!(
            package
                .authored_content_types_xml(limits)
                .unwrap()
                .as_slice(),
            package.source_content_types(limits).unwrap().bytes(),
        );
        assert_eq!(
            package
                .authored_relationships_xml(&root, limits)
                .unwrap()
                .as_slice(),
            package.source_relationships(&root, limits).unwrap().bytes(),
        );
        let small = SourceLimits::builder()
            .max_content_types_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(package.authored_content_types_xml(small).is_err());
        let small = SourceLimits::builder()
            .max_relationship_xml_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(package.authored_relationships_xml(&root, small).is_err());
        let small = SourceLimits::builder()
            .max_total_relationship_xml_events(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(package.authored_relationships_xml(&root, small).is_err());
    }

    #[test]
    fn content_types_source_token_restores_removed_override_after_reopen() {
        let source = create_minimal_docx();
        let original = OpcPackage::from_bytes(&source, SourceLimits::default()).unwrap();
        let token = original
            .source_content_types(SourceLimits::default())
            .unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();

        let mut removed = original.clone();
        assert!(removed.remove_part(&document));
        assert_eq!(removed.content_types_source_state.load(), None);
        let output = crate::PackageWriter::to_bytes(&removed).unwrap();
        let mut reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
        let current_after_removal = reopened
            .source_content_types(SourceLimits::default())
            .unwrap();
        assert_ne!(current_after_removal.bytes(), token.bytes());
        reopened
            .try_add_source_part(Box::new(BlobPart::new(
                document,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                    .to_owned(),
                original
                    .get_part(&PackURI::new("/word/document.xml").unwrap())
                    .unwrap()
                    .blob()
                    .to_vec(),
            )))
            .unwrap();
        assert_eq!(reopened.content_types_source_state.load(), None);
        let current = reopened
            .source_content_types(SourceLimits::default())
            .unwrap();
        assert!(
            reopened
                .try_replace_content_types(current.bytes(), &token, SourceLimits::default())
                .unwrap()
        );
        assert_eq!(reopened.content_types_source_state.load(), None);
        assert_eq!(
            reopened
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            token.bytes()
        );
        let restored = crate::PackageWriter::to_bytes(&reopened).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(
            &restored,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        let original_archive = soapberry_zip::office::ArchiveReader::new(
            &source,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert_eq!(
            archive.read("[Content_Types].xml").unwrap(),
            original_archive.read("[Content_Types].xml").unwrap()
        );
    }

    #[test]
    fn bounded_content_types_replacement_checks_token_before_mutation() {
        let (source, _) = create_source_with_explicit_relationship_overrides(false);
        let original = OpcPackage::from_bytes(&source, SourceLimits::default()).unwrap();
        let token = original
            .source_content_types(SourceLimits::default())
            .unwrap();
        let custom = PackURI::new("/custom/item.bin").unwrap();
        let replacement = token
            .with_part_overrides(&[(&custom, "application/octet-stream")], 4096)
            .unwrap();
        let mut package = original.clone();
        package.add_part(Box::new(BlobPart::new(
            custom,
            "application/octet-stream".to_owned(),
            b"payload".to_vec(),
        )));
        let current = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        let current_bytes = current.bytes().to_vec();
        let before_bytes = crate::PackageWriter::to_bytes(&package).unwrap();

        let under = SourceLimits::builder()
            .max_content_types_bytes(current_bytes.len())
            .unwrap()
            .build()
            .unwrap();
        let error = package
            .try_replace_content_types(&current_bytes, &replacement, under)
            .unwrap_err();
        assert!(matches!(
            error,
            OpcError::Limit {
                kind: crate::SourceLimitKind::ContentTypesBytes,
                observed: actual,
                maximum,
            } if actual == replacement.bytes().len() as u64
                && maximum == current_bytes.len() as u64
        ));
        assert_eq!(
            crate::PackageWriter::to_bytes(&package).unwrap(),
            before_bytes
        );

        let exact = SourceLimits::builder()
            .max_content_types_bytes(replacement.bytes().len())
            .unwrap()
            .max_part_bytes(replacement.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .try_replace_content_types(&current_bytes, &replacement, exact)
                .unwrap()
        );
        assert_eq!(
            package.source_content_types(exact).unwrap().bytes(),
            replacement.bytes()
        );
    }

    #[test]
    fn content_types_source_retains_existing_relationship_overrides_after_unrelated_edit() {
        for empty in [false, true] {
            let (source, content_types) = create_source_with_explicit_relationship_overrides(empty);
            let mut package = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
            assert_eq!(
                package
                    .source_content_types(SourceLimits::default())
                    .unwrap()
                    .bytes(),
                content_types
            );

            let media = PackURI::new("/media.bin").unwrap();
            package
                .get_part_mut(&media)
                .unwrap()
                .set_blob(b"changed media".to_vec());
            let output = crate::PackageWriter::to_bytes(&package).unwrap();
            let archive = soapberry_zip::office::ArchiveReader::new(
                &output,
                soapberry_zip::office::ArchiveLimits::default(),
            )
            .unwrap();
            assert_eq!(archive.read("[Content_Types].xml").unwrap(), content_types);
            assert_eq!(archive.read("media.bin").unwrap(), b"changed media");

            let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
            assert_eq!(
                reopened
                    .source_content_types(SourceLimits::default())
                    .unwrap()
                    .bytes(),
                content_types
            );
        }
    }

    #[test]
    fn content_types_source_rejects_unused_relationship_override() {
        let (source, content_types) = create_source_with_unused_relationship_override();
        let package = OpcPackage::from_bytes(&source, SourceLimits::default()).unwrap();
        let token = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        assert_ne!(token.bytes(), content_types);
        assert!(
            !token
                .bytes()
                .windows(b"document.xml.rels".len())
                .any(|window| window == b"document.xml.rels")
        );
    }

    #[test]
    fn authored_content_types_token_is_deterministic_without_source_manifest() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/item.bin").unwrap(),
            "application/octet-stream".to_owned(),
            b"payload".to_vec(),
        )));

        let first = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        let second = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(first.bytes())
                .unwrap()
                .contains("PartName=\"/custom/item.bin\"")
        );
    }

    #[test]
    fn bounded_content_types_capture_checks_retained_and_authored_quota_boundaries() {
        let package =
            OpcPackage::from_bytes(&create_minimal_docx(), SourceLimits::default()).unwrap();
        let retained = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        let exact = SourceLimits::builder()
            .max_content_types_bytes(retained.bytes().len())
            .unwrap()
            .max_part_bytes(retained.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            package.source_content_types(exact).unwrap().bytes(),
            retained.bytes()
        );
        let under = SourceLimits::builder()
            .max_content_types_bytes(retained.bytes().len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_content_types(under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::ContentTypesBytes,
                ..
            })
        ));
        let part_under = SourceLimits::builder()
            .max_content_types_bytes(retained.bytes().len())
            .unwrap()
            .max_part_bytes((retained.bytes().len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_content_types(part_under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::PartBytes,
                ..
            })
        ));
        let over = SourceLimits::builder()
            .max_content_types_bytes(retained.bytes().len() + 1)
            .unwrap()
            .max_part_bytes((retained.bytes().len() + 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(package.source_content_types(over).is_ok());

        let mut authored = OpcPackage::new();
        authored.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/item.bin").unwrap(),
            "application/octet-stream".to_owned(),
            Vec::new(),
        )));
        let authored_token = authored
            .source_content_types(SourceLimits::default())
            .unwrap();
        let authored_exact = SourceLimits::builder()
            .max_content_types_bytes(authored_token.bytes().len())
            .unwrap()
            .max_part_bytes(authored_token.bytes().len() as u64)
            .unwrap()
            .max_content_type_mappings(3)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            authored
                .source_content_types(authored_exact)
                .unwrap()
                .bytes(),
            authored_token.bytes()
        );
        let authored_events = authored_token
            .bytes()
            .iter()
            .filter(|&&byte| byte == b'<')
            .count()
            + 1;
        let authored_structure_exact = SourceLimits::builder()
            .max_events(authored_events)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            authored
                .source_content_types(authored_structure_exact)
                .is_ok()
        );
        let authored_events_under = SourceLimits::builder()
            .max_events(authored_events - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types(authored_events_under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::Events,
                observed: actual,
                maximum,
            }) if actual == authored_events as u64 && maximum == (authored_events - 1) as u64
        ));
        let authored_depth_under = SourceLimits::builder()
            .max_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types(authored_depth_under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::Depth,
                observed: 2,
                maximum: 1,
            })
        ));
        let mapping_under = SourceLimits::builder()
            .max_content_type_mappings(2)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types(mapping_under),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::ContentTypeMappings,
                observed: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn ordinary_save_keeps_explicitly_admitted_large_content_types_source() {
        let (archive, content_types) = large_content_types_archive();
        let explicit = SourceLimits::builder()
            .max_content_types_bytes(content_types.len())
            .unwrap()
            .build()
            .unwrap();
        let package = OpcPackage::from_vec(archive, explicit).unwrap();

        assert!(matches!(
            package.source_content_types(SourceLimits::default()),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::ContentTypesBytes,
                ..
            })
        ));
        let captured = package.source_content_types(explicit).unwrap();
        assert_eq!(captured.bytes(), content_types.as_slice());
        assert_eq!(
            package.source_content_types(package.read_limits()).unwrap(),
            captured
        );

        let saved = crate::PackageWriter::to_bytes(&package).unwrap();
        let saved_archive = soapberry_zip::office::ArchiveReader::new(
            &saved,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert_eq!(
            saved_archive.read("[Content_Types].xml").unwrap(),
            content_types.as_slice()
        );
    }

    #[test]
    fn bounded_content_types_checks_retained_provenance_before_fallback() {
        let (archive, content_types) = large_content_types_archive();
        let admission = SourceLimits::builder()
            .max_content_types_bytes(content_types.len())
            .unwrap()
            .build()
            .unwrap();
        let mut package = OpcPackage::from_vec(archive, admission).unwrap();
        let document = PackURI::new("/document.xml").unwrap();
        package
            .get_part_mut(&document)
            .unwrap()
            .set_content_type("application/example+xml".to_owned())
            .unwrap();

        // The invalidated package has a small canonical fallback, so the
        // caller's reduced bound is sufficient for authored output. Capture
        // still inspects the retained provenance first and must reject its
        // oversized source bytes under that same bound.
        let fallback_limit = SourceLimits::builder()
            .max_content_types_bytes(content_types.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        let saved = crate::PackageWriter::to_bytes(&package).unwrap();
        let saved_archive = soapberry_zip::office::ArchiveReader::new(
            &saved,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert!(
            saved_archive.read("[Content_Types].xml").unwrap().len()
                < fallback_limit.max_content_types_bytes()
        );
        assert!(matches!(
            package.source_content_types(fallback_limit),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::ContentTypesBytes,
                observed: actual,
                maximum,
            }) if actual == content_types.len() as u64 && maximum == (content_types.len() - 1) as u64
        ));
    }

    #[test]
    fn same_length_deferred_replacement_is_regenerated_without_hydrating_twice() {
        let target = PackURI::new("/custom/unrelated.bin").unwrap();
        let mut original_payload = Vec::with_capacity(64 * 1024);
        let mut state = 0x1234_5678_u32;
        for _ in 0..64 * 1024 {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            original_payload.push((state >> 24) as u8);
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/document.xml").unwrap(),
            crate::constants::content_type::XML.to_owned(),
            b"<document/>".to_vec(),
        )));
        package.add_part(Box::new(BlobPart::new(
            target.clone(),
            "application/octet-stream".to_owned(),
            original_payload.clone(),
        )));
        let source = crate::PackageWriter::to_bytes(&package).unwrap();
        let archive = soapberry_zip::ZipArchive::from_slice(&source).unwrap();
        let target_member = target.membername().as_bytes().to_owned();
        let mut target_compressed = 0_u64;
        let mut retained_compressed = 0_u64;
        for entry in archive.entries() {
            let entry = entry.unwrap();
            if entry.is_dir() {
                continue;
            }
            let compressed = entry.compressed_size_hint();
            if entry.file_path().as_ref() == target_member.as_slice() {
                target_compressed = compressed;
            } else {
                retained_compressed = retained_compressed.max(compressed);
            }
        }
        assert!(target_compressed > retained_compressed);

        let mut recipient = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        assert_eq!(recipient.deferred_decode_counters(), Some((0, 0)));
        let original_len = original_payload.len() as u64;
        let mut replacement = original_payload;
        replacement[0] ^= 0xff;
        let replacement_for_assertion = replacement.clone();
        recipient
            .replace_part_payloads(vec![(target, replacement)])
            .unwrap();
        let after_replacement = recipient.deferred_decode_counters();
        assert_eq!(after_replacement, Some((1, original_len)));

        let final_metrics = recipient
            .retained_archive_metrics(SourceLimits::default())
            .unwrap()
            .unwrap();
        let limits = SourceLimits::builder()
            .max_archive_compressed_bytes(retained_compressed)
            .unwrap()
            .max_archive_metadata_bytes(final_metrics.metadata_bytes())
            .unwrap()
            .build()
            .unwrap();
        assert!(recipient.retained_archive_metrics(limits).is_ok());
        assert_eq!(recipient.deferred_decode_counters(), after_replacement);
        recipient.read_limits = limits;
        let output = crate::PackageWriter::to_bytes(&recipient).unwrap();
        let reopened = OpcPackage::from_vec(output, SourceLimits::default()).unwrap();
        assert_eq!(
            reopened
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .unwrap()
                .blob(),
            replacement_for_assertion.as_slice()
        );
        assert_eq!(recipient.deferred_decode_counters(), after_replacement);
    }

    #[test]
    fn batch_source_token_replacement_preserves_absent_relationship_member() {
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<document/>".to_vec(),
        )));
        package
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:test".to_owned(),
                "https://example.test/target".to_owned(),
                "rId1".to_owned(),
                crate::TargetMode::External,
            )
            .unwrap();

        let content_types = package
            .source_content_types(SourceLimits::default())
            .unwrap();
        let expected_relationships = package
            .source_relationships(&owner, SourceLimits::default())
            .unwrap();
        assert!(expected_relationships.member_present());

        let mut absent_source = OpcPackage::new();
        absent_source.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<document/>".to_vec(),
        )));
        let replacement_relationships = absent_source
            .source_relationships(&owner, SourceLimits::default())
            .unwrap();
        assert!(!replacement_relationships.member_present());

        // The batch must honor the package's admission limits while reading
        // its preconditions, even when supplied tokens were captured earlier
        // under more generous limits. Neither refusal may publish the edit.
        let original_limits = package.read_limits;
        let source_bytes = crate::PackageWriter::to_bytes(&package).unwrap();
        for (limits, resource) in [
            (
                SourceLimits::builder()
                    .max_content_types_bytes(content_types.bytes().len() - 1)
                    .unwrap()
                    .build()
                    .unwrap(),
                crate::SourceLimitKind::ContentTypesBytes,
            ),
            (
                SourceLimits::builder()
                    .max_relationship_xml_bytes(expected_relationships.bytes().len() - 1)
                    .unwrap()
                    .build()
                    .unwrap(),
                crate::SourceLimitKind::RelationshipXmlBytes,
            ),
        ] {
            package.read_limits = limits;
            assert!(matches!(
                package.try_add_parts_with_source_tokens(
                    content_types.bytes(),
                    &content_types,
                    &expected_relationships,
                    &replacement_relationships,
                    Vec::new(),
                ),
                Err(OpcError::Limit { kind: actual, .. }) if actual == resource
            ));
            package.read_limits = original_limits;
            assert_eq!(
                crate::PackageWriter::to_bytes(&package).unwrap(),
                source_bytes
            );
        }

        package
            .try_add_parts_with_source_tokens(
                content_types.bytes(),
                &content_types,
                &expected_relationships,
                &replacement_relationships,
                Vec::new(),
            )
            .unwrap();

        let after = package
            .source_relationships(&owner, SourceLimits::default())
            .unwrap();
        assert!(!after.member_present());
        assert_eq!(after.bytes(), replacement_relationships.bytes());

        let output = crate::PackageWriter::to_bytes(&package).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        let relationship_member = owner.rels_uri().unwrap().membername().to_owned();
        assert!(!archive.file_names().any(|name| name == relationship_member));
        let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
        assert!(
            !reopened
                .source_relationships(&owner, SourceLimits::default())
                .unwrap()
                .member_present()
        );
    }

    #[test]
    fn batch_source_tokens_re_admit_broad_metadata_under_recipient_limits() {
        let archive = create_minimal_docx();
        // The producer fixture deliberately adds a relationship target larger
        // than the default 4 KiB target ceiling. Open it under an actually
        // broader owner policy so the source token can be created; the
        // recipient below remains tight and must reject it atomically.
        let producer_limits = SourceLimits::builder()
            .max_relationship_target_bytes(8192)
            .unwrap()
            .build()
            .unwrap();
        let producer = OpcPackage::from_bytes(&archive, producer_limits).unwrap();
        let root = PackURI::new(PACKAGE_URI).unwrap();
        let broad_content_types = producer
            .source_content_types(producer.read_limits())
            .unwrap();
        let expected_relationships = producer
            .source_relationships(&root, producer.read_limits())
            .unwrap();
        let broad_relationships = expected_relationships
            .with_relationship(
                "urn:test:large",
                &format!("https://example.test/{}", "x".repeat(4096)),
                "rId-large",
                crate::TargetMode::External,
                usize::MAX,
            )
            .unwrap();
        let largest_existing_relationship = producer
            .source_relationships_xml
            .values()
            .map(|source| source.bytes.len())
            .max()
            .unwrap();
        let recipient_limits = SourceLimits::builder()
            .max_relationship_xml_bytes(largest_existing_relationship)
            .unwrap()
            // Keep the target ceiling above the deliberately broad token so
            // this regression isolates the recipient XML-byte refusal.
            .max_relationship_target_bytes(8192)
            .unwrap()
            .build()
            .unwrap();
        let mut recipient = OpcPackage::from_bytes(&archive, recipient_limits).unwrap();
        let expected_content_types = recipient
            .source_content_types(recipient.read_limits())
            .unwrap();
        let current_relationships = recipient
            .source_relationships(&root, recipient.read_limits())
            .unwrap();
        let before = crate::PackageWriter::to_bytes(&recipient).unwrap();

        let error = recipient
            .try_add_parts_with_source_tokens(
                expected_content_types.bytes(),
                &broad_content_types,
                &current_relationships,
                &broad_relationships,
                Vec::new(),
            )
            .unwrap_err();
        assert!(matches!(
            error,
            OpcError::Limit {
                kind: crate::SourceLimitKind::RelationshipXmlBytes,
                observed,
                maximum,
            } if observed == broad_relationships.bytes().len() as u64
                && maximum == largest_existing_relationship as u64
        ));
        assert_eq!(crate::PackageWriter::to_bytes(&recipient).unwrap(), before);
        assert_eq!(
            recipient
                .source_relationships(&root, recipient.read_limits())
                .unwrap(),
            current_relationships
        );
    }

    #[test]
    fn batch_source_token_replacement_honors_relationship_limit_atomically() {
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut original = OpcPackage::new();
        original.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<document/>".to_vec(),
        )));
        let source = crate::PackageWriter::to_bytes(&original).unwrap();
        let mut package = OpcPackage::from_vec(source.clone(), SourceLimits::default()).unwrap();
        let content_types = package.source_content_types(package.read_limits()).unwrap();
        let expected_relationships = package
            .source_relationships(&owner, package.read_limits())
            .unwrap();
        assert!(!expected_relationships.member_present());

        let mut replacement_source = OpcPackage::new();
        replacement_source.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<document/>".to_vec(),
        )));
        replacement_source
            .get_part_mut(&owner)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:test:replacement".to_owned(),
                "https://example.test/target".to_owned(),
                "rId1".to_owned(),
                crate::TargetMode::External,
            )
            .unwrap();
        let replacement_relationships = replacement_source
            .source_relationships(&owner, replacement_source.read_limits())
            .unwrap();
        assert!(replacement_relationships.member_present());
        assert!(replacement_relationships.bytes().len() > 1);

        let tight = SourceLimits::builder()
            .max_relationship_xml_bytes(replacement_relationships.bytes().len() - 1)
            .unwrap()
            .build()
            .unwrap();
        package.read_limits = tight;
        assert!(matches!(
            package.try_add_parts_with_source_tokens(
                content_types.bytes(),
                &content_types,
                &expected_relationships,
                &replacement_relationships,
                Vec::new(),
            ),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::RelationshipXmlBytes,
                ..
            })
        ));
        assert_eq!(crate::PackageWriter::to_bytes(&package).unwrap(), source);
        assert!(package.is_unmodified_owned_source());
        assert_eq!(
            package.exact_source_shared().unwrap().as_slice(),
            source.as_slice()
        );
    }

    #[test]
    fn font_embedding_policy_has_only_three_valid_states() {
        let mut package = OpcPackage::new();
        assert_eq!(package.save_options().fonts, FontEmbedding::None);
        package.with_fonts(FontEmbedding::Subset);
        assert_eq!(package.save_options().fonts, FontEmbedding::Subset);
        package.with_fonts(FontEmbedding::Full);
        assert_eq!(package.save_options().fonts, FontEmbedding::Full);
    }

    #[test]
    fn relationship_provenance_uses_empty_sentinel_and_exact_owned_bytes() {
        let empty_relationships = Relationships::new(PACKAGE_URI.to_owned());
        let empty = PreservedRelationshipsXml::from_relationships(&empty_relationships).unwrap();
        assert!(matches!(empty, PreservedRelationshipsXml::Empty));
        assert_eq!(empty.as_bytes(), crate::rel::EMPTY_RELATIONSHIPS_XML);

        let mut relationships = Relationships::new(PACKAGE_URI.to_owned());
        relationships
            .try_add_relationship(
                "urn:test".to_owned(),
                "/custom/data.bin".to_owned(),
                "rId1".to_owned(),
                crate::TargetMode::Internal,
            )
            .unwrap();
        let expected = relationships.try_to_xml_bytes().unwrap();
        let owned = PreservedRelationshipsXml::from_relationships(&relationships).unwrap();
        assert!(matches!(&owned, PreservedRelationshipsXml::Owned(_)));
        assert_eq!(owned.as_bytes(), expected.as_slice());
    }

    #[test]
    fn moves_owned_archive_into_package_reader() {
        let pkg = OpcPackage::from_vec(create_minimal_docx(), SourceLimits::default()).unwrap();
        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn owned_constructors_preserve_mixed_storage_and_exact_source() {
        let source = with_eocd_comment(create_mixed_minimal_docx(), b"mixed storage");
        let file = NamedTempFile::new().unwrap();
        fs::write(file.path(), &source).unwrap();

        let from_vec = OpcPackage::from_vec(source.clone(), SourceLimits::default()).unwrap();
        let from_reader =
            OpcPackage::from_reader(Cursor::new(source.clone()), SourceLimits::default()).unwrap();
        let from_path = OpcPackage::open(file.path(), SourceLimits::default()).unwrap();
        let partname = PackURI::new("/word/document.xml").unwrap();

        for package in [&from_vec, &from_reader, &from_path] {
            assert_eq!(package.get_part(&partname).unwrap().blob(),
                       b"<?xml version=\"1.0\"?>\n<document xmlns=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n    <body><p><t>Test</t></p></body>\n</document>");
            assert_eq!(crate::PackageWriter::to_bytes(package).unwrap(), source);
        }
    }

    #[test]
    fn owned_constructors_check_input_limit_before_zip_validation() {
        let source = b"four".to_vec();
        let file = NamedTempFile::new().unwrap();
        fs::write(file.path(), &source).unwrap();
        let limits = SourceLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();

        for result in [
            OpcPackage::from_vec(source.clone(), limits),
            OpcPackage::from_reader(Cursor::new(source.clone()), limits),
            OpcPackage::open(file.path(), limits),
        ] {
            assert!(matches!(
                result,
                Err(OpcError::Limit {
                    kind: crate::SourceLimitKind::InputBytes,
                    observed: 4,
                    maximum: 3,
                })
            ));
        }
    }

    #[test]
    fn owned_constructors_report_malformed_zip_after_bounded_read() {
        let source = b"not an OPC ZIP".to_vec();
        let file = NamedTempFile::new().unwrap();
        fs::write(file.path(), &source).unwrap();

        for result in [
            OpcPackage::from_vec(source.clone(), SourceLimits::default()),
            OpcPackage::from_reader(Cursor::new(source.clone()), SourceLimits::default()),
            OpcPackage::open(file.path(), SourceLimits::default()),
        ] {
            assert!(matches!(result, Err(OpcError::ZipError(_))));
        }
    }

    #[test]
    fn every_durability_level_publishes_the_exact_source_and_keeps_its_authorization() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"exact source 0761");
        let package = OpcPackage::from_vec(bytes.clone(), SourceLimits::default())
            .expect("open owned package");
        let digest = package.exact_source_sha256();
        assert!(digest.is_some());
        let directory = tempfile::tempdir().unwrap();

        let reference = directory.path().join("reference.docx");
        package.save(&reference).unwrap();
        assert_eq!(fs::read(&reference).unwrap(), bytes);

        for durability in [
            litchi_core::Durability::Full,
            litchi_core::Durability::FileOnly,
            litchi_core::Durability::NoSync,
        ] {
            let destination = directory
                .path()
                .join(format!("{}.docx", durability.as_str()));
            fs::write(&destination, b"old destination").unwrap();
            package
                .save_with_durability(&destination, durability)
                .unwrap();
            // The exact no-op publication is untouched by the level, and the
            // level, being no save preference, revokes nothing.
            assert_eq!(fs::read(&destination).unwrap(), bytes, "{durability:?}");
            assert!(package.is_unmodified_owned_source(), "{durability:?}");
            assert_eq!(package.exact_source_sha256(), digest);
        }
        assert_eq!(fs::read_dir(directory.path()).unwrap().count(), 4);
    }

    #[test]
    fn clone_shares_owned_source_but_revocation_is_independent() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"exact source");
        let package =
            OpcPackage::from_vec(bytes, SourceLimits::default()).expect("open owned package");
        let mut clone = package.clone();

        assert!(package.is_unmodified_owned_source());
        assert!(clone.is_unmodified_owned_source());
        assert!(!OpcPackage::new().is_unmodified_owned_source());
        assert!(
            !OpcPackage::from_bytes(&create_minimal_docx(), SourceLimits::default())
                .unwrap()
                .is_unmodified_owned_source()
        );

        assert!(Arc::ptr_eq(
            package
                .source_archive
                .as_ref()
                .expect("source authorized")
                .bytes(),
            clone
                .source_archive
                .as_ref()
                .expect("clone source authorized")
                .bytes()
        ));

        let unchanged_options = clone.save_options().clone();
        clone.set_save_options(unchanged_options);
        assert!(!clone.is_unmodified_owned_source());
        assert!(package.is_unmodified_owned_source());
        assert!(!clone.exact_source_authorized);
        assert!(clone.source_archive.is_some());
        assert!(clone.preservation.is_some());
        assert!(package.exact_source_authorized);
    }

    fn sha256_of(bytes: &[u8]) -> [u8; 32] {
        use sha2::{Digest as _, Sha256};
        Sha256::digest(bytes).into()
    }

    /// The exact-source digest is the digest of exactly the bytes the package
    /// streams, on every owned and shared ingress, and is absent otherwise.
    #[test]
    fn exact_source_sha256_is_the_digest_of_the_published_bytes() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"exact source digest");
        let expected = sha256_of(&bytes);
        let file = NamedTempFile::new().unwrap();
        fs::write(file.path(), &bytes).unwrap();
        let donor = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        let shared = Arc::new(bytes.clone());
        let packages = [
            OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap(),
            OpcPackage::from_reader(Cursor::new(bytes.clone()), SourceLimits::default()).unwrap(),
            OpcPackage::open(file.path(), SourceLimits::default()).unwrap(),
            OpcPackage::from_vec_reusing_payloads(bytes.clone(), SourceLimits::default(), &donor)
                .unwrap(),
            OpcPackage::from_shared_vec_reusing_payloads(
                Arc::clone(&shared),
                SourceLimits::default(),
                &donor,
            )
            .unwrap(),
            OpcPackage::from_shared_vec(Arc::clone(&shared), SourceLimits::default()).unwrap(),
        ];
        for package in &packages {
            assert_eq!(package.exact_source_sha256(), Some(expected));
            assert_eq!(package.exact_source_len(), Some(bytes.len()));
            let mut streamed = Vec::new();
            package.write_to(&mut streamed).unwrap();
            assert_eq!(sha256_of(&streamed), expected);
            assert_eq!(streamed, bytes);
        }

        assert_eq!(OpcPackage::new().exact_source_sha256(), None);
        assert_eq!(OpcPackage::new().exact_source_len(), None);
        assert_eq!(
            OpcPackage::from_bytes(&bytes, SourceLimits::default())
                .unwrap()
                .exact_source_sha256(),
            None,
            "borrowed ingress retains no exact source"
        );
    }

    /// Clones share one memo, filled at most once from the archive they share;
    /// a revoked clone answers nothing, and a second ingress of equal bytes
    /// owns a memo of its own.
    #[test]
    fn exact_source_sha256_memo_is_bound_to_its_archive() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"memo binding");
        let package = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        let source = package.source_archive.as_ref().unwrap();
        assert!(!source.sha256_is_memoized(), "ingress hashes nothing");

        let mut clone = package.clone();
        let cloned_source = clone.source_archive.as_ref().unwrap();
        assert!(source.shares_memo_with(cloned_source));
        assert_eq!(clone.exact_source_sha256(), Some(sha256_of(&bytes)));
        assert!(
            source.sha256_is_memoized(),
            "a clone's digest is the original's: they share one archive"
        );

        clone.set_save_options(SaveOptions::default());
        assert_eq!(clone.exact_source_sha256(), None, "an edit revokes it");
        assert_eq!(clone.exact_source_len(), None, "and the length with it");
        assert_eq!(package.exact_source_sha256(), Some(sha256_of(&bytes)));

        let again = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        let again_source = again.source_archive.as_ref().unwrap();
        assert!(!again_source.shares_memo_with(source));
        assert!(!again_source.sha256_is_memoized());
        assert_eq!(again.exact_source_sha256(), package.exact_source_sha256());

        let other = with_eocd_comment(create_minimal_docx(), b"other bytes");
        let different = OpcPackage::from_vec(other.clone(), SourceLimits::default()).unwrap();
        assert_eq!(different.exact_source_sha256(), Some(sha256_of(&other)));
        assert_ne!(
            different.exact_source_sha256(),
            package.exact_source_sha256()
        );
    }

    /// Shared ingress keeps the caller's allocation instead of a copy, reads
    /// and refuses exactly as the owned ingress does, and donates the same
    /// payload allocations.
    #[test]
    fn shared_ingress_matches_owned_ingress_without_copying() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"shared ingress");
        let donor = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        let shared = Arc::new(bytes.clone());

        let owned =
            OpcPackage::from_vec_reusing_payloads(bytes.clone(), SourceLimits::default(), &donor)
                .unwrap();
        let reused = OpcPackage::from_shared_vec_reusing_payloads(
            Arc::clone(&shared),
            SourceLimits::default(),
            &donor,
        )
        .unwrap();
        let deferred =
            OpcPackage::from_shared_vec(Arc::clone(&shared), SourceLimits::default()).unwrap();
        for package in [&reused, &deferred] {
            assert!(package.is_unmodified_owned_source());
            assert!(Arc::ptr_eq(
                &package.exact_source_shared().unwrap(),
                &shared
            ));
            assert_eq!(package.source_read_limits(), Some(SourceLimits::default()));
            assert_eq!(package.part_count(), owned.part_count());
        }
        for metadata in owned.iter_parts() {
            let name = metadata.partname();
            let expected = owned.get_part(name).unwrap();
            let donated = reused.get_part(name).unwrap();
            let lazy = deferred.get_part(name).unwrap();
            assert_eq!(expected.content_type(), donated.content_type());
            assert_eq!(expected.content_type(), lazy.content_type());
            assert_eq!(expected.blob(), donated.blob());
            assert_eq!(expected.blob(), lazy.blob());
            assert_eq!(
                Arc::ptr_eq(
                    &expected.blob_arc(),
                    &donor.get_part(name).unwrap().blob_arc()
                ),
                Arc::ptr_eq(
                    &donated.blob_arc(),
                    &donor.get_part(name).unwrap().blob_arc()
                ),
                "shared ingress donates exactly what owned ingress donates"
            );
        }

        let tight = SourceLimits::builder()
            .max_input_bytes(8)
            .unwrap()
            .build()
            .unwrap();
        let handles = Arc::strong_count(&shared);
        let owned_refusal = OpcPackage::from_vec(bytes.clone(), tight)
            .err()
            .map(|error| error.to_string());
        let shared_refusal = OpcPackage::from_shared_vec(Arc::clone(&shared), tight)
            .err()
            .map(|error| error.to_string());
        assert!(owned_refusal.is_some());
        assert_eq!(owned_refusal, shared_refusal);
        let reused_refusal =
            OpcPackage::from_shared_vec_reusing_payloads(Arc::clone(&shared), tight, &donor)
                .err()
                .map(|error| error.to_string());
        assert_eq!(owned_refusal, reused_refusal);
        assert_eq!(
            Arc::strong_count(&shared),
            handles,
            "a refused shared ingress keeps no handle"
        );
    }

    #[test]
    fn admitted_reader_handoff_rejects_a_reader_from_another_source() {
        let source = Arc::new(create_minimal_docx());
        let foreign = Arc::new(source.as_ref().clone());
        let reader = PhysPkgReader::new(source.as_slice(), SourceLimits::default()).unwrap();

        let result = OpcPackage::from_shared_vec_with_admitted_reader(foreign, reader);
        assert!(matches!(
            result,
            Err(OpcError::SourceArtifactMismatch {
                artifact: "shared OPC source",
                field: "physical reader bytes",
            })
        ));
    }

    #[test]
    fn admitted_reader_handoff_retains_the_reader_policy() {
        let source = Arc::new(create_minimal_docx());
        let limits = SourceLimits::builder()
            .max_input_bytes(source.len() as u64)
            .unwrap()
            .build()
            .unwrap();
        let reader = PhysPkgReader::new(source.as_slice(), limits).unwrap();

        let package =
            OpcPackage::from_shared_vec_with_admitted_reader(Arc::clone(&source), reader).unwrap();
        assert_eq!(package.read_limits(), limits);
        assert_eq!(package.source_read_limits(), Some(limits));
        assert!(Arc::ptr_eq(
            &package.exact_source_shared().unwrap(),
            &source
        ));
    }

    #[test]
    fn mutable_api_entries_revoke_owned_source_even_on_failure_or_noop() {
        let source = with_eocd_comment(create_minimal_docx(), b"exact source");
        let base =
            OpcPackage::from_vec(source, SourceLimits::default()).expect("open owned package");
        let missing = PackURI::new("/missing.xml").expect("valid URI");

        let mut package = base.clone();
        assert!(package.get_part_mut(&missing).is_err());
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        assert!(!package.remove_part(&missing));
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        let duplicate = package.main_document_part().unwrap().partname().clone();
        let error = package.try_add_part(Box::new(BlobPart::new(
            duplicate,
            "application/xml".to_owned(),
            Vec::new(),
        )));
        assert!(error.is_err());
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        package.unsign();
        assert!(!package.exact_source_authorized);
    }

    #[test]
    fn batch_payload_replacement_preflights_all_targets_and_empty_is_exact_noop() {
        let source = create_two_part_source();
        let base = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let second = PackURI::new("/word/second.xml").unwrap();
        let before_source = base.exact_source_shared().unwrap();
        let before_document = base.get_part(&document).unwrap().blob().to_vec();
        let before_second = base.get_part(&second).unwrap().blob().to_vec();

        let mut empty = base.clone();
        empty.replace_part_payloads(Vec::new()).unwrap();
        assert!(empty.is_unmodified_owned_source());
        assert!(Arc::ptr_eq(
            &before_source,
            &empty.exact_source_shared().unwrap()
        ));

        let mut package = base;
        let missing = PackURI::new("/word/missing.xml").unwrap();
        let error = package
            .replace_part_payloads(vec![
                (
                    PackURI::new("/WORD/DOCUMENT.XML").unwrap(),
                    b"<changed/>".to_vec(),
                ),
                (missing, b"<missing/>".to_vec()),
            ])
            .unwrap_err();
        assert!(matches!(error, OpcError::PartNotFound(_)));
        assert_eq!(package.get_part(&document).unwrap().blob(), before_document);
        assert_eq!(package.get_part(&second).unwrap().blob(), before_second);
        assert!(package.is_unmodified_owned_source());
        assert!(Arc::ptr_eq(
            &before_source,
            &package.exact_source_shared().unwrap()
        ));
    }

    #[test]
    fn batch_payload_replacement_rejects_case_alias_duplicates_before_mutation() {
        let source = create_two_part_source();
        let mut package = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let before = package.get_part(&document).unwrap().blob().to_vec();

        let error = package
            .replace_part_payloads(vec![
                (document.clone(), b"<first/>".to_vec()),
                (
                    PackURI::new("/WORD/DOCUMENT.XML").unwrap(),
                    b"<last/>".to_vec(),
                ),
            ])
            .unwrap_err();
        assert!(matches!(error, OpcError::DuplicatePartPayload(_)));
        assert_eq!(package.get_part(&document).unwrap().blob(), before);
        assert!(package.is_unmodified_owned_source());
    }

    #[test]
    fn batch_payload_replacement_preflights_custom_decode_refusal() {
        let good = PackURI::new("/custom/good.xml").unwrap();
        let failing = PackURI::new("/custom/failing.xml").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            good.clone(),
            crate::constants::content_type::XML.to_owned(),
            b"<good/>".to_vec(),
        )));
        package.add_part(Box::new(DecodeRefusingPart {
            partname: failing.clone(),
            payload: Arc::new(b"<failing/>".to_vec()),
            rels: Relationships::new(failing.as_str().to_owned()),
        }));
        let before = package.get_part(&good).unwrap().blob().to_vec();

        let error = package
            .replace_part_payloads(vec![
                (good.clone(), b"<changed/>".to_vec()),
                (failing, b"<replacement/>".to_vec()),
            ])
            .unwrap_err();
        assert!(matches!(error, OpcError::ZipError(message) if message == "test decode refusal"));
        assert_eq!(package.get_part(&good).unwrap().blob(), before);
    }

    #[test]
    fn batch_payload_replacement_refuses_custom_parts_before_any_setter() {
        let first = PackURI::new("/custom/first.xml").unwrap();
        let second = PackURI::new("/custom/second.xml").unwrap();
        let first_calls = Arc::new(AtomicUsize::new(0));
        let second_calls = Arc::new(AtomicUsize::new(0));
        let mut package = OpcPackage::new();
        package.add_part(Box::new(StatefulDecodePart {
            partname: first.clone(),
            payload: Arc::new(b"<first/>".to_vec()),
            rels: Relationships::new(first.as_str().to_owned()),
            ensure_calls: Arc::clone(&first_calls),
            fail_on_call: usize::MAX,
        }));
        package.add_part(Box::new(StatefulDecodePart {
            partname: second.clone(),
            payload: Arc::new(b"<second/>".to_vec()),
            rels: Relationships::new(second.as_str().to_owned()),
            ensure_calls: Arc::clone(&second_calls),
            fail_on_call: 2,
        }));

        let error = package
            .replace_part_payloads(vec![
                (first.clone(), b"<changed-first/>".to_vec()),
                (second.clone(), b"<changed-second/>".to_vec()),
            ])
            .unwrap_err();
        assert!(matches!(error, OpcError::UnsupportedCustomPartPayload));
        assert_eq!(package.parts.get(&first).unwrap().blob(), b"<first/>");
        assert_eq!(package.parts.get(&second).unwrap().blob(), b"<second/>");
        assert_eq!(first_calls.load(Ordering::SeqCst), 1);
        assert_eq!(second_calls.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn batch_payload_replacement_revokes_source_without_bypassing_xml_provenance() {
        let source = create_two_part_source();
        let mut package = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let replacement = b"<document><body><p/></body></document>".to_vec();

        package
            .replace_part_payloads(vec![(document.clone(), replacement)])
            .unwrap();
        assert!(!package.is_unmodified_owned_source());
        assert!(!package.holds_original_source_xml(package.get_part(&document).unwrap()));
    }

    #[test]
    fn batch_payload_replacement_preserves_injected_signature_policy() {
        let source = create_two_part_source();
        let mut package = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();
        let second = PackURI::new("/word/second.xml").unwrap();
        package
            .get_part_mut(&document)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:vendor:signature-reference".to_owned(),
                "/_xmlsignatures/orphan.xml".to_owned(),
                "rId-signature".to_owned(),
                crate::TargetMode::Internal,
            )
            .unwrap();
        assert!(package.is_signed());
        assert!(package.requires_signature_edit_policy());

        package
            .replace_part_payloads(vec![(second, b"<changed/>".to_vec())])
            .unwrap();
        assert!(package.is_signed());
        assert!(package.requires_signature_edit_policy());
        assert!(matches!(
            crate::PackageWriter::to_bytes(&package),
            Err(OpcError::SignedSourceRequiresExplicitPolicy)
        ));
        assert!(package.signature_policy_required);
    }

    #[test]
    fn relationship_and_option_mutable_apis_revoke_owned_source() {
        let source = with_eocd_comment(create_minimal_docx(), b"exact source");
        let base =
            OpcPackage::from_vec(source, SourceLimits::default()).expect("open owned package");

        let mut package = base.clone();
        let options = package.save_options().clone();
        package.set_save_options(options);
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        package.with_fonts(FontEmbedding::None);
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        let _relationships = package.rels_mut();
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        let _relationships = package.relationships_mut();
        assert!(!package.exact_source_authorized);

        let mut package = base.clone();
        package.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        assert!(!package.exact_source_authorized);

        let mut package = base;
        package.relate_to_external("https://example.com", "urn:example");
        assert!(!package.exact_source_authorized);
    }

    #[test]
    fn signature_detection_covers_orphans_and_relationship_targets() {
        let mut orphan = OpcPackage::new();
        orphan.add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/orphan.xml").unwrap(),
            "application/octet-stream".to_owned(),
            Vec::new(),
        )));
        assert!(orphan.is_signed());
        orphan.unsign();
        assert!(!orphan.is_signed());

        let mut owner = BlobPart::new(
            PackURI::new("/custom/owner.bin").unwrap(),
            "application/octet-stream".to_owned(),
            Vec::new(),
        );
        Part::relate_to(
            &mut owner,
            "../_xmlsignatures/orphan.xml",
            "urn:vendor:signature-reference",
        );
        let mut targeted = OpcPackage::new();
        targeted.add_part(Box::new(owner));
        assert!(targeted.is_signed());
        targeted.unsign();
        assert!(!targeted.is_signed());
    }

    #[test]
    fn removing_origin_relationship_does_not_bypass_signature_detection() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            crate::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )));
        package.relate_to(
            "_xmlsignatures/origin.sigs",
            relationship_type::DIGITAL_SIGNATURE_ORIGIN,
        );
        assert!(package.is_signed());
        package.rels_mut().remove("rId1");
        assert!(package.is_signed());
        package.unsign();
        assert!(!package.is_signed());
    }

    #[test]
    fn part_relationship_mutation_retains_signature_policy_tracking() {
        let mut package =
            OpcPackage::from_bytes(&create_minimal_docx(), SourceLimits::default()).unwrap();
        let partname = package.main_document_part().unwrap().partname().clone();
        package
            .get_part_mut(&partname)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:vendor:signature-reference".to_owned(),
                "/_xmlsignatures/orphan.xml".to_owned(),
                "rId-signature".to_owned(),
                crate::TargetMode::Internal,
            )
            .unwrap();

        assert!(package.is_signed());
        assert!(package.requires_signature_edit_policy());
    }

    #[test]
    fn candidate_signature_policy_checks_source_provenance_and_disposition() {
        let mut authored = OpcPackage::new();
        authored.add_part(Box::new(BlobPart::new(
            PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
            crate::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            Vec::new(),
        )));
        let bytes = crate::PackageWriter::to_bytes(&authored).unwrap();
        let original = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        assert!(
            original
                .clone()
                .validate_signature_edit_from(&original)
                .is_ok()
        );
        let mut candidate = original.clone();
        candidate.remove_part(&PackURI::new("/_xmlsignatures/origin.sigs").unwrap());
        assert!(!candidate.is_signed());
        assert!(candidate.validate_signature_edit_from(&original).is_err());
        candidate.unsign();
        assert!(candidate.validate_signature_edit_from(&original).is_ok());

        let mut replacement = OpcPackage::from_vec(bytes.clone(), SourceLimits::default()).unwrap();
        replacement.unsign();
        assert!(replacement.validate_signature_edit_from(&original).is_err());
        assert!(
            OpcPackage::new()
                .validate_signature_edit_from(&original)
                .is_err()
        );

        let mut borrowed = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
        let mut candidate = borrowed.clone();
        candidate.unsign();
        assert!(candidate.validate_signature_edit_from(&borrowed).is_err());
        borrowed.unsign();
        assert!(candidate.validate_signature_edit_from(&borrowed).is_ok());
        assert!(
            OpcPackage::new()
                .validate_signature_edit_from(&borrowed)
                .is_ok()
        );
    }

    #[test]
    fn bounded_package_constructors_reject_oversized_input() {
        let limits = SourceLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OpcPackage::from_bytes(b"four", limits),
            Err(OpcError::Limit {
                kind: crate::SourceLimitKind::InputBytes,
                observed: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn package_retains_ingress_limits_through_clone_and_edits() {
        let bytes = create_minimal_docx();
        let limits = SourceLimits::builder()
            .max_input_bytes(bytes.len() as u64)
            .unwrap()
            .max_part_bytes(8192)
            .unwrap()
            .build()
            .unwrap();
        let borrowed = OpcPackage::from_bytes(&bytes, limits).unwrap();
        let owned = OpcPackage::from_vec(bytes.clone(), limits).unwrap();
        let streamed = OpcPackage::from_reader(Cursor::new(bytes), limits).unwrap();
        for package in [borrowed, owned, streamed] {
            assert_eq!(package.read_limits(), limits);
            let mut edited = package.clone();
            edited
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new("/new.bin").unwrap(),
                    "application/octet-stream".to_owned(),
                    b"new".to_vec(),
                )))
                .unwrap();
            assert_eq!(edited.read_limits(), limits);
            assert_eq!(package.read_limits(), limits);
        }
        assert_eq!(OpcPackage::new().read_limits(), SourceLimits::default());
    }

    #[test]
    fn test_main_document_part() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor, SourceLimits::default()).unwrap();

        let main_part = pkg.main_document_part().unwrap();
        assert_eq!(
            main_part.content_type(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        );
    }

    #[test]
    fn resolves_one_internal_strict_main_document_part() {
        let mut package = OpcPackage::new();
        let uri = PackURI::new("/custom/main.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            uri.clone(),
            "application/xml".to_string(),
            Vec::new(),
        )));
        package.relate_to("custom/main.xml", relationship_type::STRICT_OFFICE_DOCUMENT);

        assert_eq!(package.main_document_part().unwrap().partname(), &uri);

        package.relate_to("other.xml", relationship_type::OFFICE_DOCUMENT);
        assert!(package.main_document_part().is_err());
    }

    #[test]
    fn remembered_main_document_part_follows_relationship_changes() {
        let mut package = OpcPackage::new();
        for name in ["/word/document.xml", "/word/other.xml"] {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new(name).unwrap(),
                "application/xml".to_string(),
                Vec::new(),
            )));
        }
        let r_id = package.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        let first = PackURI::new("/word/document.xml").unwrap();
        assert_eq!(package.main_document_part().unwrap().partname(), &first);
        // Remembered: a clone and a repeated call resolve the same part.
        assert_eq!(package.main_document_part().unwrap().partname(), &first);
        let clone = package.clone();
        assert_eq!(clone.main_document_part().unwrap().partname(), &first);

        // A retarget, a removal and a replacement are each seen at once.
        package
            .rels_mut()
            .retarget(&r_id, "word/other.xml".to_owned())
            .unwrap();
        assert_eq!(
            package.main_document_part().unwrap().partname().as_str(),
            "/word/other.xml"
        );
        package.rels_mut().remove(&r_id);
        assert!(package.main_document_part().is_err());
        package.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        assert_eq!(package.main_document_part().unwrap().partname(), &first);
        assert_eq!(clone.main_document_part().unwrap().partname(), &first);
    }

    #[test]
    fn try_add_part_rejects_conflicts_without_replacing_the_original() {
        let mut package = OpcPackage::new();
        let original_uri = PackURI::new("/word/document.xml").unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                original_uri.clone(),
                "application/xml".to_string(),
                b"original".to_vec(),
            )))
            .unwrap();

        for (candidate, expected) in [
            ("/word/document.xml", PartNameConflict::Duplicate),
            ("/WORD/DOCUMENT.XML", PartNameConflict::Equivalent),
            ("/word/document.xml/image.gif", PartNameConflict::Derived),
        ] {
            let error = package
                .try_add_part(Box::new(BlobPart::new(
                    PackURI::new(candidate).unwrap(),
                    "application/octet-stream".to_string(),
                    b"replacement".to_vec(),
                )))
                .unwrap_err();
            assert!(matches!(
                (expected, error),
                (PartNameConflict::Duplicate, OpcError::DuplicatePartName(_))
                    | (
                        PartNameConflict::Equivalent,
                        OpcError::EquivalentPartNames { .. }
                    )
                    | (PartNameConflict::Derived, OpcError::DerivedPartNames { .. })
            ));
            assert_eq!(package.part_count(), 1);
            assert_eq!(package.get_part(&original_uri).unwrap().blob(), b"original");
        }
    }

    #[test]
    fn materialized_part_batch_preflights_all_name_conflicts_without_mutation() {
        let make_part = |name: &str| -> (Box<dyn Part + Send + Sync>, bool) {
            (
                Box::new(BlobPart::new(
                    PackURI::new(name).unwrap(),
                    "application/xml".to_owned(),
                    b"<part/>".to_vec(),
                )),
                true,
            )
        };
        for (names, expected) in [
            (
                vec!["/word/document.xml", "/word/document.xml"],
                PartNameConflict::Duplicate,
            ),
            (
                vec!["/word/document.xml", "/WORD/DOCUMENT.XML"],
                PartNameConflict::Equivalent,
            ),
            (
                vec!["/word/document.xml", "/word/document.xml/image.gif"],
                PartNameConflict::Derived,
            ),
            (
                vec!["/word/document.xml/image.gif", "/word/document.xml"],
                PartNameConflict::Derived,
            ),
        ] {
            let mut package = OpcPackage::new();
            let parts = names.into_iter().map(make_part).collect();
            let error = package
                .try_add_materialized_parts(parts, || Ok(()))
                .unwrap_err();
            assert!(matches!(
                (expected, error),
                (PartNameConflict::Duplicate, OpcError::DuplicatePartName(_))
                    | (
                        PartNameConflict::Equivalent,
                        OpcError::EquivalentPartNames { .. }
                    )
                    | (PartNameConflict::Derived, OpcError::DerivedPartNames { .. })
            ));
            assert_eq!(package.part_count(), 0);
            assert!(!package.exact_source_authorized);
            assert!(!package.signature_graph_tracked);
            assert!(!package.signature_policy_authorized);
            assert!(!package.signature_policy_required);
            assert!(package.source_xml_parts.is_empty());
        }

        for (candidate, expected) in [
            ("/word/document.xml", PartNameConflict::Duplicate),
            ("/WORD/DOCUMENT.XML", PartNameConflict::Equivalent),
            ("/word/document.xml/image.gif", PartNameConflict::Derived),
        ] {
            let mut package =
                OpcPackage::from_vec(create_minimal_docx(), SourceLimits::default()).unwrap();
            let exact_source_authorized = package.exact_source_authorized;
            let signature_graph_tracked = package.signature_graph_tracked;
            let signature_policy_authorized = package.signature_policy_authorized;
            let signature_policy_required = package.signature_policy_required;
            let source_xml_count = package.source_xml_parts.len();
            let error = package
                .try_add_materialized_parts(vec![make_part(candidate)], || Ok(()))
                .unwrap_err();
            assert!(matches!(
                (expected, error),
                (PartNameConflict::Duplicate, OpcError::DuplicatePartName(_))
                    | (
                        PartNameConflict::Equivalent,
                        OpcError::EquivalentPartNames { .. }
                    )
                    | (PartNameConflict::Derived, OpcError::DerivedPartNames { .. })
            ));
            assert_eq!(package.part_count(), 1);
            assert_eq!(package.exact_source_authorized, exact_source_authorized);
            assert_eq!(package.signature_graph_tracked, signature_graph_tracked);
            assert_eq!(
                package.signature_policy_authorized,
                signature_policy_authorized
            );
            assert_eq!(package.signature_policy_required, signature_policy_required);
            assert_eq!(package.source_xml_parts.len(), source_xml_count);
        }
    }

    #[test]
    fn materialized_part_batch_retains_source_xml_provenance_only_for_source_parts() {
        let source_name = PackURI::new("/word/document.xml").unwrap();
        let authored_name = PackURI::new("/word/styles.xml").unwrap();
        let mut package = OpcPackage::new();
        package
            .try_add_materialized_parts(
                vec![
                    (
                        Box::new(BlobPart::new(
                            source_name.clone(),
                            "application/xml".to_owned(),
                            b"<document/>".to_vec(),
                        )),
                        true,
                    ),
                    (
                        Box::new(BlobPart::new(
                            authored_name.clone(),
                            "application/xml".to_owned(),
                            b"<styles/>".to_vec(),
                        )),
                        false,
                    ),
                ],
                || Ok(()),
            )
            .unwrap();

        assert!(package.holds_original_source_xml(package.get_part(&source_name).unwrap()));
        assert!(!package.holds_original_source_xml(package.get_part(&authored_name).unwrap()));
    }

    #[test]
    fn package_clone_shares_clean_payloads_and_detaches_mutation() {
        let uri = PackURI::new("/custom/data.bin").unwrap();
        let mut source = OpcPackage::new();
        source
            .try_add_part(Box::new(BlobPart::new(
                uri.clone(),
                "application/octet-stream".to_string(),
                b"source".to_vec(),
            )))
            .unwrap();

        let source_blob = source.get_part(&uri).unwrap().blob_arc();
        let mut edited = source.clone();
        let edited_blob = edited.get_part(&uri).unwrap().blob_arc();
        assert!(Arc::ptr_eq(&source_blob, &edited_blob));

        let replacement = Arc::new(b"edited".to_vec());
        edited
            .get_part_mut(&uri)
            .unwrap()
            .set_blob_shared(Arc::clone(&replacement));
        assert_eq!(source.get_part(&uri).unwrap().blob(), b"source");
        assert_eq!(edited.get_part(&uri).unwrap().blob(), b"edited");
        assert!(Arc::ptr_eq(
            &edited.get_part(&uri).unwrap().blob_arc(),
            &replacement
        ));
    }

    #[test]
    fn source_xml_provenance_requires_the_ingress_allocation() {
        let uri = PackURI::new("/custom/source.xml").unwrap();
        let source_xml = b"<root>\n  <child/>\n</root>".to_vec();
        let mut authored = OpcPackage::new();
        authored.add_part(Box::new(BlobPart::new(
            uri.clone(),
            "application/xml".to_owned(),
            source_xml,
        )));

        let bytes = crate::pkgwriter::PackageWriter::to_bytes(&authored).unwrap();
        let mut reopened = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
        assert!(reopened.holds_original_source_xml(reopened.get_part(&uri).unwrap()));

        // Equal bytes in a newly allocated payload are still authored from the
        // writer's point of view. A byte comparison would incorrectly retain
        // the source proof here; allocation identity must revoke it.
        let equal_copy = reopened.get_part(&uri).unwrap().blob().to_vec();
        reopened.get_part_mut(&uri).unwrap().set_blob(equal_copy);
        assert!(!reopened.holds_original_source_xml(reopened.get_part(&uri).unwrap()));
    }
}
