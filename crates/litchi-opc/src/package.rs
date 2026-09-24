//! Objects that implement reading and writing OPC packages.
//!
//! This module provides the main `OpcPackage` type, which represents an Open Packaging
//! Convention package in memory. It manages parts, relationships, and provides
//! high-level operations for working with office documents.

use crate::constants::relationship_type;
use crate::content_type::ContentTypeMap;
use crate::error::{OpcError, Result};
use crate::execution::OpenSession;
use crate::limits::ReadLimits;
use crate::members::NonPartMember;
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
use std::sync::{Arc, OnceLock};

mod content_types;
mod relationships;
pub use content_types::{ContentTypeEdit, ContentTypesEditPlan};
pub use relationships::{
    CanonicalRelationshipsPlan, OwnedRelationships, RelationshipEdit, RelationshipSourcePlan,
    RelationshipsEditPlan,
};

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
    limits: ReadLimits,
) -> Result<()> {
    limits.check(
        crate::ReadResource::RelationshipsPerPart,
        relationships.len() as u64,
        limits.max_relationships_per_part() as u64,
    )?;
    limits.check(
        crate::ReadResource::TotalRelationships,
        relationships.len() as u64,
        limits.max_total_relationships() as u64,
    )?;
    for relationship in relationships.iter() {
        limits.check(
            crate::ReadResource::XmlAttributeBytes,
            relationship.r_id().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::ReadResource::XmlAttributeBytes,
            relationship.reltype().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::ReadResource::XmlAttributeBytes,
            relationship.target_ref().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        limits.check(
            crate::ReadResource::RelationshipTargetBytes,
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
    limits: ReadLimits,
) -> Result<()> {
    limits.check(
        crate::ReadResource::XmlAttributeBytes,
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
                crate::ReadResource::XmlAttributeBytes,
                actual as u64,
                limits.max_xml_attribute_bytes() as u64,
            )?;
        }
        if relationship.target_mode() == RelationshipTargetMode::External {
            let actual = "TargetMode".len() + b"External".len();
            limits.check(
                crate::ReadResource::XmlAttributeBytes,
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
    limits: ReadLimits,
) -> Result<()> {
    let events = relationships.len().checked_add(4).ok_or_else(|| {
        OpcError::InvalidRelationship("canonical relationship XML event count overflows".to_owned())
    })?;
    limits.check(
        crate::ReadResource::XmlEvents,
        events as u64,
        limits.max_xml_events() as u64,
    )?;
    limits.check(
        crate::ReadResource::TotalRelationshipXmlEvents,
        events as u64,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    let depth = if relationships.is_empty() { 1 } else { 2 };
    limits.check(
        crate::ReadResource::XmlDepth,
        depth,
        limits.max_xml_depth() as u64,
    )?;
    limits.check(
        crate::ReadResource::TotalRelationshipXmlBytes,
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

    /// Remove exact part and relationship-part overrides while preserving all
    /// surrounding manifest bytes. Defaults, comments, ordering, and unrelated
    /// overrides remain source-backed in the returned token.
    pub fn without_parts(&self, parts: &[PackURI], max_output_bytes: usize) -> Result<Self> {
        let xml = content_types::without_part_overrides(&self.xml, parts, max_output_bytes)?;
        if std::ptr::eq(xml.bytes(), self.xml.bytes()) {
            return Ok(self.clone());
        }
        let binding = Arc::new(ContentTypeMap::from_xml(&xml.bytes, ReadLimits::default())?);
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
        let binding = Arc::new(ContentTypeMap::from_xml(&xml.bytes, ReadLimits::default())?);
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
    read_limits: ReadLimits,
    /// Package-level relationships
    rels: Relationships,

    /// All parts in the package, indexed by partname
    /// Using Box<dyn Part + Send + Sync> for trait objects to allow different part types
    /// `PackURI` keys avoid string allocations compared to String keys
    parts: HashMap<PackURI, Box<dyn Part + Send + Sync>>,

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
    source_limits: ReadLimits,

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

    /// Save preferences
    save_options: SaveOptions,
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
            .field("save_options", &self.save_options)
            .finish()
    }
}

impl OpcPackage {
    /// Create a new empty OPC package.
    #[must_use]
    pub fn new() -> Self {
        Self {
            read_limits: ReadLimits::default(),
            rels: Relationships::new(PACKAGE_URI.to_string()),
            parts: HashMap::new(),
            source_xml_parts: HashMap::new(),
            source_relationships_xml: HashMap::new(),
            source_content_types_xml: None,
            source_content_types: None,
            source_archive: None,
            source_limits: ReadLimits::default(),
            transfer_index: None,
            exact_source_authorized: false,
            source_ingress: false,
            signature_graph_tracked: false,
            signature_policy_authorized: false,
            signature_policy_required: false,
            signature_api_authored: false,
            preservation: None,
            non_part_members: Vec::new(),
            save_options: SaveOptions::default(),
        }
    }

    /// Return the policy used to read this package.
    ///
    /// Newly authored packages use the default policy. Cloning or editing a
    /// package retains its captured policy so format-level editors can apply
    /// the caller's bounds before staging new payloads. This accessor does
    /// not itself validate edits or change publication behavior.
    #[must_use]
    pub const fn read_limits(&self) -> ReadLimits {
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

    /// Replace the reader-reported ZIP members that are not OPC parts.
    ///
    /// This is crate-private because only package ingress paths can establish
    /// the classification; mutable package callers cannot manufacture it.
    pub(crate) fn set_non_part_members(&mut self, members: Vec<NonPartMember>) {
        self.non_part_members = members;
    }

    pub(crate) fn source_relationships_member_present(&self, partname: &PackURI) -> bool {
        self.source_relationships_xml.contains_key(partname)
    }

    fn check_source_content_types_limits(
        &self,
        name: &PackURI,
        bytes: &[u8],
        source: &ContentTypeMap,
        limits: ReadLimits,
    ) -> Result<()> {
        limits.check(
            crate::ReadResource::ContentTypesBytes,
            bytes.len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            crate::ReadResource::PartBytes,
            bytes.len() as u64,
            limits.max_part_bytes(),
        )?;
        limits.check(
            crate::ReadResource::Parts,
            self.parts.len() as u64,
            limits.max_parts() as u64,
        )?;
        limits.check(
            crate::ReadResource::ContentTypeMappings,
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
        self.source_content_types_matches_current_parts(map)
            .map(|matches| matches.then_some((bytes, map)))
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

    /// Capture the current `[Content_Types].xml` publication view.
    ///
    /// Compatible admitted source bytes are retained exactly. After an edit
    /// that invalidates source coverage, the view is deterministic authored
    /// XML for the current part/type set. Newly authored packages use the
    /// same deterministic view. Callers cannot construct the token from
    /// arbitrary XML, and replacement is accepted only for the current
    /// expected view.
    pub fn source_content_types(&self) -> Result<OwnedContentTypes> {
        self.source_content_types_with_limits(ReadLimits::default())
    }

    /// Capture the current `[Content_Types].xml` publication view under an
    /// explicit bounded read policy. Retained source bytes and declarations
    /// are checked before compatibility indexing; authored fallback uses the
    /// same limits for its two-pass size and mapping preflight.
    pub fn source_content_types_with_limits(
        &self,
        limits: ReadLimits,
    ) -> Result<OwnedContentTypes> {
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
                    crate::ReadResource::ContentTypesBytes,
                    bytes.len() as u64,
                    limits.max_content_types_bytes() as u64,
                )?;
                limits.check(
                    crate::ReadResource::PartBytes,
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

    /// Replace an exact source-manifest snapshot.
    ///
    /// The caller owns the format-level dependency closure. Publication only
    /// reuses the token once the current part/type coverage agrees with its
    /// parsed declarations, so restoring the token before re-adding a removed
    /// part remains safe and deterministic.
    pub fn try_replace_content_types(
        &mut self,
        expected: &[u8],
        replacement: &OwnedContentTypes,
    ) -> Result<bool> {
        self.try_replace_content_types_with_limits(expected, replacement, ReadLimits::default())
    }

    /// Replace an exact source-manifest snapshot under an explicit bounded
    /// read policy.
    ///
    /// The current manifest is checked before replacement validation or any
    /// package mutation. A changed signed package is refused before parsing
    /// or retaining replacement metadata, and the replacement's XML is
    /// admitted under the same limits used for the current source check.
    pub fn try_replace_content_types_with_limits(
        &mut self,
        expected: &[u8],
        replacement: &OwnedContentTypes,
        limits: ReadLimits,
    ) -> Result<bool> {
        let current = self.source_content_types_with_limits(limits)?;
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
        if current.bytes() == replacement.bytes() {
            return Ok(false);
        }
        if self.is_signed() || self.requires_signature_edit_policy() {
            return Err(OpcError::SignedSourceRequiresExplicitPolicy);
        }

        limits.check(
            crate::ReadResource::ContentTypesBytes,
            replacement.bytes().len() as u64,
            limits.max_content_types_bytes() as u64,
        )?;
        limits.check(
            crate::ReadResource::ContentTypeMappings,
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
        self.revoke_exact_source();
        self.source_content_types_xml = Some(Arc::clone(&replacement.xml.bytes));
        self.source_content_types = Some(Arc::clone(&replacement.binding));
        Ok(true)
    }

    pub(crate) fn source_relationships_xml(
        &self,
        owner: &PackURI,
        relationships: &Relationships,
    ) -> Result<Option<&[u8]>> {
        let binding = RelationshipBinding::from_relationships(relationships)?;
        Ok(self
            .source_relationships_xml
            .get(owner)
            .filter(|source| source.binding == binding)
            .map(|source| source.bytes.as_slice()))
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
    /// For such a package [`Self::to_stream`] writes exactly the retained
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
    /// authorization is intact: the number of bytes [`Self::to_stream`]
    /// writes for such a package.
    ///
    /// Returns `None` exactly when [`Self::exact_source_sha256`] does. It reads
    /// the length without taking a handle to the archive or hashing it, so a
    /// caller can check a bound before asking for the digest.
    #[must_use]
    pub fn exact_source_len(&self) -> Option<usize> {
        self.exact_source().map(<[u8]>::len)
    }

    /// The read limits this package's owned source archive was admitted
    /// under, or `None` when the package retains no owned source archive.
    ///
    /// A caller that re-admits the bytes this package publishes, for example
    /// to decide something from those bytes alone, can hold them to the
    /// policy the package itself was opened under (change 0742).
    #[must_use]
    pub fn source_read_limits(&self) -> Option<ReadLimits> {
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
    /// use litchi_opc::package::OpcPackage;
    ///
    /// let pkg = OpcPackage::open("document.docx").unwrap();
    /// ```
    ///
    /// # Errors
    /// Returns an error if the file cannot be read or is not a valid OPC package.
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        Self::open_with_limits(path, ReadLimits::default())
    }

    /// Open an OPC package from a file with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the file cannot be read, violates `limits`, or is not
    /// a valid OPC package.
    pub fn open_with_limits<P: AsRef<Path>>(path: P, limits: ReadLimits) -> Result<Self> {
        let data = read_owned_path_with_limits(path, limits)?;
        Self::from_owned_bytes_with_limits(data, limits)
    }

    /// Load an OPC package from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    ///
    /// # Errors
    /// Returns an error if the archive cannot be read or is not a valid OPC package.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_limits(reader, ReadLimits::default())
    }

    /// Load an OPC package from a reader with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the archive cannot be read, violates `limits`, or is
    /// not a valid OPC package.
    pub fn from_reader_with_limits<R: Read>(reader: R, limits: ReadLimits) -> Result<Self> {
        let data = read_limited(reader, limits)?;
        Self::from_owned_bytes_with_limits(data, limits)
    }

    /// Move an owned ZIP archive into the package reader.
    ///
    /// This avoids copying the archive buffer before parts are decompressed.
    ///
    /// # Errors
    /// Returns an error if the archive is not a valid OPC package.
    pub fn from_vec(data: Vec<u8>) -> Result<Self> {
        Self::from_vec_with_limits(data, ReadLimits::default())
    }

    /// Move an owned ZIP archive into the package reader with explicit limits.
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC package.
    pub fn from_vec_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
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
        limits: ReadLimits,
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
        limits: ReadLimits,
        donor: &Self,
    ) -> Result<Self> {
        let mut package = {
            let phys_reader = PhysPkgReader::new_with_limits(data.as_slice(), limits)?;
            let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
            Self::unmarshal_with_payload_donor(pkg_reader, Some(donor))?
        };
        package.authorize_shared_owned_source(data, limits);
        Ok(package)
    }

    /// [`Self::from_vec_with_limits`] over an archive allocation the caller
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
    pub fn from_shared_vec_with_limits(data: Arc<Vec<u8>>, limits: ReadLimits) -> Result<Self> {
        Self::from_shared_bytes_with_limits(data, limits)
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
        limits: ReadLimits,
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
    /// Returns an error if the archive is not a valid OPC package.
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        Self::from_bytes_with_limits(data, ReadLimits::default())
    }

    /// Load an OPC package from a byte slice with explicit resource limits.
    ///
    /// # Errors
    /// Returns an error if the archive violates `limits` or is not a valid OPC package.
    pub fn from_bytes_with_limits(data: &[u8], limits: ReadLimits) -> Result<Self> {
        let phys_reader = PhysPkgReader::new_with_limits(data, limits)?;
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
        limits: ReadLimits,
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

        // Create all parts - move data instead of cloning
        for spart in sparts {
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
                part.rels_mut().try_add_relationship(
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
            package.rels.try_add_relationship(
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
        package.source_xml_parts = source_xml_parts;
        package.source_ingress = true;
        package.signature_graph_tracked = package.is_signed();
        package.signature_policy_required = package.signature_graph_tracked;
        Ok(package)
    }

    fn retain_relationships_xml(
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
    /// equality test with this proof; the preservation planner's own
    /// `source_blob_retained` keeps its byte comparison, because there an
    /// equal payload publishes equal bytes either way.
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
        crate::OwnedXmlPart::capture(name.clone(), part.content_type().into(), bytes)
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
        // A deferred payload reports its own refusal rather than comparing as
        // an empty member (ADR 0030).
        current.ensure_payload()?;
        if current.content_type() != replacement.content_type || current.blob() != expected {
            return Err(OpcError::XmlError(
                "stale owned XML part replacement".into(),
            ));
        }
        // These bytes are recorded as the part's source provenance below, and
        // provenance exempts a payload from the writer's publication audit
        // (change 0665). They came from the caller, so audit them now, before
        // any mutation, and hand the part the proof (change 0754).
        let verified = verified_source_payload(&replacement.name, Arc::clone(&replacement.bytes))?;
        self.source_xml_parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "owned XML provenance",
                source,
            })?;
        let part = self.get_part_mut(&replacement.name)?;
        part.set_blob_verified(verified);
        let payload = part.payload_handle().payload().clone();
        self.source_xml_parts.insert(replacement.name, payload);
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
        let content_type = self.get_part(partname)?.content_type().to_owned();
        let token = crate::OwnedXmlPart::capture_with_limits(
            partname.clone(),
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
        self.try_add_source_part(Box::new(crate::BlobPart::new_shared(
            source.name,
            source.content_type,
            source.bytes,
        )))
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
        let partname = rel.target_partname()?;
        self.get_part(&partname)
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
        self.parts
            .get_mut(&key)
            .map(|b| {
                let part: &mut dyn Part = &mut **b;
                part
            })
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
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
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Add a new part to the package.
    ///
    /// # Arguments
    /// * `part` - The part to add
    pub fn add_part(&mut self, part: Box<dyn Part + Send + Sync>) {
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        let partname = part.partname().clone();
        self.source_xml_parts.remove(&partname);
        self.source_relationships_xml.remove(&partname);
        self.parts.insert(partname, part);
    }

    /// Try to add a part without replacing an existing or ambiguous part name.
    ///
    /// # Errors
    /// Returns an error if the part's partname duplicates or conflicts with an
    /// existing part name; the existing part is left untouched.
    pub fn try_add_part(&mut self, part: Box<dyn Part + Send + Sync>) -> Result<()> {
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        let partname = part.partname().clone();
        self.validate_new_part_name(&partname)?;
        self.parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC package parts",
                source,
            })?;
        self.source_relationships_xml.remove(&partname);
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
        let current_content_types = self.source_content_types_with_limits(self.read_limits)?;
        if current_content_types.bytes() != expected_content_types {
            return Err(OpcError::InvalidContentTypesManifest(
                "stale content-types source replacement".to_owned(),
            ));
        }
        let current_relationships = self
            .source_relationships_with_limits(expected_relationships.owner(), self.read_limits)?;
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

        let replacement_relationships_parsed = PackageReader::parse_owned_relationships(
            replacement_relationships.bytes(),
            replacement_relationships.owner(),
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
            let declared = replacement_content_types
                .binding
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
            let declared = replacement_content_types
                .binding
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
        self.signature_graph_tracked = true;
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
        self.source_content_types = Some(Arc::clone(&replacement_content_types.binding));
        Ok(())
    }

    /// Add a source-materialized part while retaining its exact XML bytes for
    /// the publication audit. Source-backed conversion uses this narrow seam
    /// so unchanged opaque XML remains publishable without reparsing or
    /// normalizing it. The source blob is shared with the inserted part and
    /// the metadata map grows through a fallible reservation.
    pub(crate) fn try_add_source_part(&mut self, part: Box<dyn Part + Send + Sync>) -> Result<()> {
        self.revoke_exact_source();
        self.signature_graph_tracked = true;
        let partname = part.partname().clone();
        self.validate_new_part_name(&partname)?;
        self.parts
            .try_reserve(1)
            .map_err(|source| OpcError::Allocation {
                resource: "OPC source-materialized parts",
                source,
            })?;
        let source_blob =
            if xml_minifier::audit::package::is_xml_part(partname.as_str(), part.content_type()) {
                self.source_xml_parts
                    .try_reserve(1)
                    .map_err(|source| OpcError::Allocation {
                        resource: "OPC source-preserved XML parts",
                        source,
                    })?;
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
        self.source_xml_parts.remove(partname);
        self.source_relationships_xml.remove(partname);
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

    /// Get the number of parts in the package.
    #[must_use]
    pub fn part_count(&self) -> usize {
        self.parts.len()
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
    /// pkg.to_stream(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    /// Returns an error if the package cannot be serialized or written to the
    /// stream; the stream may be left incomplete.
    pub fn to_stream<W: Write>(&self, writer: W) -> Result<()> {
        crate::pkgwriter::PackageWriter::write_to_stream(writer, self)
    }

    pub(crate) fn strip_signature_graph(&mut self) {
        self.revoke_exact_source();
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
    /// (`from_shared_vec_with_limits`, `from_shared_vec_reusing_payloads`)
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

    fn from_owned_bytes_with_limits(data: Vec<u8>, limits: ReadLimits) -> Result<Self> {
        Self::from_shared_bytes_with_limits(Arc::new(data), limits)
    }

    fn from_shared_bytes_with_limits(data: Arc<Vec<u8>>, limits: ReadLimits) -> Result<Self> {
        let mut package = {
            let phys_reader = PhysPkgReader::new_with_limits(data.as_slice(), limits)?;
            let pkg_reader = PackageReader::from_phys_reader_deferred(&phys_reader, &data)?;
            Self::unmarshal(pkg_reader)?
        };
        package.authorize_shared_owned_source(data, limits);
        Ok(package)
    }

    pub(crate) fn from_bytes_with_open_session(
        data: &[u8],
        limits: ReadLimits,
        session: &OpenSession,
    ) -> Result<Self> {
        session.check()?;
        let phys_reader = PhysPkgReader::new_with_limits(data, limits)?;
        session.charge_input(data.len() as u64)?;
        let pkg_reader = PackageReader::from_phys_reader_with_session(&phys_reader, session)?;
        Self::unmarshal(pkg_reader)
    }

    pub(crate) fn from_vec_with_open_session(
        data: Vec<u8>,
        limits: ReadLimits,
        session: &OpenSession,
    ) -> Result<Self> {
        session.check()?;
        let input_bytes = data.len() as u64;
        let mut package = {
            let phys_reader = PhysPkgReader::new_with_limits(&data, limits)?;
            session.charge_input(input_bytes)?;
            let pkg_reader = PackageReader::from_phys_reader_with_session(&phys_reader, session)?;
            Self::unmarshal(pkg_reader)?
        };
        package.authorize_owned_source(data, limits);
        Ok(package)
    }

    fn authorize_owned_source(&mut self, source: Vec<u8>, limits: ReadLimits) {
        self.authorize_shared_owned_source(Arc::new(source), limits);
    }

    /// Retain an owned source archive that deferred payloads already share.
    fn authorize_shared_owned_source(&mut self, source: Arc<Vec<u8>>, limits: ReadLimits) {
        let preservation = PreservationProvenance::from_package(source.as_slice(), self);
        if let Some(preservation) = preservation.as_ref() {
            self.bind_relationship_captures(preservation);
        }
        self.preservation = preservation.map(Arc::new);
        self.source_archive = Some(retained_archive::RetainedArchive::new(source));
        self.source_limits = limits;
        self.transfer_index = Some(Arc::new(OnceLock::new()));
        self.exact_source_authorized = true;
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
    xml_minifier::audit::VerifiedSource::verify(bytes, xml_minifier::audit::Limits::default())
        .map_err(|source| OpcError::XmlPublication {
            part: name.to_string(),
            source,
        })
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
    use tempfile::NamedTempFile;

    fn create_minimal_docx() -> Vec<u8> {
        create_minimal_docx_with_stored_document(false)
    }

    fn create_mixed_minimal_docx() -> Vec<u8> {
        create_minimal_docx_with_stored_document(true)
    }

    fn create_minimal_docx_with_stored_document(stored_document: bool) -> Vec<u8> {
        let mut writer = StreamingArchiveWriter::new();

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

    fn large_content_types_archive() -> (Vec<u8>, Vec<u8>) {
        let mut content_types = Vec::new();
        content_types.extend_from_slice(
            br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><!--"#,
        );
        content_types.resize(
            content_types.len() + ReadLimits::default().max_content_types_bytes() + 1,
            b'x',
        );
        content_types.extend_from_slice(
            br#"--><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/document.xml" ContentType="application/xml"/></Types>"#,
        );

        let mut writer = StreamingArchiveWriter::new();
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
        let mut writer = StreamingArchiveWriter::new();
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
        let mut writer = StreamingArchiveWriter::new();
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
        let pkg = OpcPackage::from_reader(cursor).unwrap();

        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn relationship_source_at_raw_limit_does_not_use_expanding_canonical_cache() {
        let target = format!("https://example.test/{}", ">".repeat(2048));
        let relationships = format!(
            r#"<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="{target}" TargetMode="External"/></Relationships>"#
        );
        let mut writer = StreamingArchiveWriter::new();
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
        let limits = ReadLimits::builder()
            .max_relationship_xml_bytes(relationships.len())
            .unwrap()
            .max_total_relationship_xml_bytes(relationships.len())
            .unwrap()
            .build()
            .unwrap();

        let package = OpcPackage::from_bytes_with_limits(&source, limits).unwrap();
        let owner = PackURI::new(PACKAGE_URI).unwrap();
        let token = package.source_relationships(&owner).unwrap();
        assert_eq!(token.bytes(), relationships.as_bytes());
    }

    #[test]
    fn content_types_source_token_restores_removed_override_after_reopen() {
        let source = create_minimal_docx();
        let original = OpcPackage::from_bytes(&source).unwrap();
        let token = original.source_content_types().unwrap();
        let document = PackURI::new("/word/document.xml").unwrap();

        let mut removed = original.clone();
        assert!(removed.remove_part(&document));
        let output = crate::PackageWriter::to_bytes(&removed).unwrap();
        let mut reopened = OpcPackage::from_bytes(&output).unwrap();
        let current_after_removal = reopened.source_content_types().unwrap();
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
        let current = reopened.source_content_types().unwrap();
        assert!(
            reopened
                .try_replace_content_types(current.bytes(), &token)
                .unwrap()
        );
        let restored = crate::PackageWriter::to_bytes(&reopened).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
        let original_archive = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
        assert_eq!(
            archive.read("[Content_Types].xml").unwrap(),
            original_archive.read("[Content_Types].xml").unwrap()
        );
    }

    #[test]
    fn bounded_content_types_replacement_checks_token_before_mutation() {
        let (source, _) = create_source_with_explicit_relationship_overrides(false);
        let original = OpcPackage::from_bytes(&source).unwrap();
        let token = original.source_content_types().unwrap();
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
        let current = package.source_content_types().unwrap();
        let current_bytes = current.bytes().to_vec();
        let before_bytes = crate::PackageWriter::to_bytes(&package).unwrap();

        let under = ReadLimits::builder()
            .max_content_types_bytes(current_bytes.len())
            .unwrap()
            .build()
            .unwrap();
        let error = package
            .try_replace_content_types_with_limits(&current_bytes, &replacement, under)
            .unwrap_err();
        assert!(matches!(
            error,
            OpcError::ReadLimit {
                resource: crate::ReadResource::ContentTypesBytes,
                actual,
                maximum,
            } if actual == replacement.bytes().len() as u64
                && maximum == current_bytes.len() as u64
        ));
        assert_eq!(
            crate::PackageWriter::to_bytes(&package).unwrap(),
            before_bytes
        );

        let exact = ReadLimits::builder()
            .max_content_types_bytes(replacement.bytes().len())
            .unwrap()
            .max_part_bytes(replacement.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            package
                .try_replace_content_types_with_limits(&current_bytes, &replacement, exact)
                .unwrap()
        );
        assert_eq!(
            package
                .source_content_types_with_limits(exact)
                .unwrap()
                .bytes(),
            replacement.bytes()
        );
    }

    #[test]
    fn content_types_source_retains_existing_relationship_overrides_after_unrelated_edit() {
        for empty in [false, true] {
            let (source, content_types) = create_source_with_explicit_relationship_overrides(empty);
            let mut package = OpcPackage::from_vec(source).unwrap();
            assert_eq!(
                package.source_content_types().unwrap().bytes(),
                content_types
            );

            let media = PackURI::new("/media.bin").unwrap();
            package
                .get_part_mut(&media)
                .unwrap()
                .set_blob(b"changed media".to_vec());
            let output = crate::PackageWriter::to_bytes(&package).unwrap();
            let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
            assert_eq!(archive.read("[Content_Types].xml").unwrap(), content_types);
            assert_eq!(archive.read("media.bin").unwrap(), b"changed media");

            let reopened = OpcPackage::from_bytes(&output).unwrap();
            assert_eq!(
                reopened.source_content_types().unwrap().bytes(),
                content_types
            );
        }
    }

    #[test]
    fn content_types_source_rejects_unused_relationship_override() {
        let (source, content_types) = create_source_with_unused_relationship_override();
        let package = OpcPackage::from_bytes(&source).unwrap();
        let token = package.source_content_types().unwrap();
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

        let first = package.source_content_types().unwrap();
        let second = package.source_content_types().unwrap();
        assert_eq!(first, second);
        assert!(
            std::str::from_utf8(first.bytes())
                .unwrap()
                .contains("PartName=\"/custom/item.bin\"")
        );
    }

    #[test]
    fn bounded_content_types_capture_checks_retained_and_authored_quota_boundaries() {
        let package = OpcPackage::from_bytes(&create_minimal_docx()).unwrap();
        let retained = package.source_content_types().unwrap();
        let exact = ReadLimits::builder()
            .max_content_types_bytes(retained.bytes().len())
            .unwrap()
            .max_part_bytes(retained.bytes().len() as u64)
            .unwrap()
            .build()
            .unwrap();
        assert_eq!(
            package
                .source_content_types_with_limits(exact)
                .unwrap()
                .bytes(),
            retained.bytes()
        );
        let under = ReadLimits::builder()
            .max_content_types_bytes(retained.bytes().len() - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_content_types_with_limits(under),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::ContentTypesBytes,
                ..
            })
        ));
        let part_under = ReadLimits::builder()
            .max_content_types_bytes(retained.bytes().len())
            .unwrap()
            .max_part_bytes((retained.bytes().len() - 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            package.source_content_types_with_limits(part_under),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::PartBytes,
                ..
            })
        ));
        let over = ReadLimits::builder()
            .max_content_types_bytes(retained.bytes().len() + 1)
            .unwrap()
            .max_part_bytes((retained.bytes().len() + 1) as u64)
            .unwrap()
            .build()
            .unwrap();
        assert!(package.source_content_types_with_limits(over).is_ok());

        let mut authored = OpcPackage::new();
        authored.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/item.bin").unwrap(),
            "application/octet-stream".to_owned(),
            Vec::new(),
        )));
        let authored_token = authored.source_content_types().unwrap();
        let authored_exact = ReadLimits::builder()
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
                .source_content_types_with_limits(authored_exact)
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
        let authored_structure_exact = ReadLimits::builder()
            .max_xml_events(authored_events)
            .unwrap()
            .build()
            .unwrap();
        assert!(
            authored
                .source_content_types_with_limits(authored_structure_exact)
                .is_ok()
        );
        let authored_events_under = ReadLimits::builder()
            .max_xml_events(authored_events - 1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types_with_limits(authored_events_under),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::XmlEvents,
                actual,
                maximum,
            }) if actual == authored_events as u64 && maximum == (authored_events - 1) as u64
        ));
        let authored_depth_under = ReadLimits::builder()
            .max_xml_depth(1)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types_with_limits(authored_depth_under),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::XmlDepth,
                actual: 2,
                maximum: 1,
            })
        ));
        let mapping_under = ReadLimits::builder()
            .max_content_type_mappings(2)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            authored.source_content_types_with_limits(mapping_under),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::ContentTypeMappings,
                actual: 3,
                maximum: 2,
            })
        ));
    }

    #[test]
    fn ordinary_save_keeps_explicitly_admitted_large_content_types_source() {
        let (archive, content_types) = large_content_types_archive();
        let explicit = ReadLimits::builder()
            .max_content_types_bytes(content_types.len())
            .unwrap()
            .build()
            .unwrap();
        let package = OpcPackage::from_vec_with_limits(archive, explicit).unwrap();

        assert!(matches!(
            package.source_content_types(),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::ContentTypesBytes,
                ..
            })
        ));
        let captured = package.source_content_types_with_limits(explicit).unwrap();
        assert_eq!(captured.bytes(), content_types.as_slice());

        let saved = crate::PackageWriter::to_bytes(&package).unwrap();
        let saved_archive = soapberry_zip::office::ArchiveReader::new(&saved).unwrap();
        assert_eq!(
            saved_archive.read("[Content_Types].xml").unwrap(),
            content_types.as_slice()
        );
    }

    #[test]
    fn bounded_content_types_checks_retained_provenance_before_fallback() {
        let (archive, content_types) = large_content_types_archive();
        let admission = ReadLimits::builder()
            .max_content_types_bytes(content_types.len())
            .unwrap()
            .build()
            .unwrap();
        let mut package = OpcPackage::from_vec_with_limits(archive, admission).unwrap();
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
        let fallback_limit = ReadLimits::builder()
            .max_content_types_bytes(content_types.len() - 1)
            .unwrap()
            .build()
            .unwrap();
        let saved = crate::PackageWriter::to_bytes(&package).unwrap();
        let saved_archive = soapberry_zip::office::ArchiveReader::new(&saved).unwrap();
        assert!(
            saved_archive.read("[Content_Types].xml").unwrap().len()
                < fallback_limit.max_content_types_bytes()
        );
        assert!(matches!(
            package.source_content_types_with_limits(fallback_limit),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::ContentTypesBytes,
                actual,
                maximum,
            }) if actual == content_types.len() as u64 && maximum == (content_types.len() - 1) as u64
        ));
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

        let content_types = package.source_content_types().unwrap();
        let expected_relationships = package.source_relationships(&owner).unwrap();
        assert!(expected_relationships.member_present());

        let mut absent_source = OpcPackage::new();
        absent_source.add_part(Box::new(BlobPart::new(
            owner.clone(),
            "application/xml".to_owned(),
            b"<document/>".to_vec(),
        )));
        let replacement_relationships = absent_source.source_relationships(&owner).unwrap();
        assert!(!replacement_relationships.member_present());

        // The batch must honor the package's admission limits while reading
        // its preconditions, even when supplied tokens were captured earlier
        // under more generous limits. Neither refusal may publish the edit.
        let original_limits = package.read_limits;
        let source_bytes = crate::PackageWriter::to_bytes(&package).unwrap();
        for (limits, resource) in [
            (
                ReadLimits::builder()
                    .max_content_types_bytes(content_types.bytes().len() - 1)
                    .unwrap()
                    .build()
                    .unwrap(),
                crate::ReadResource::ContentTypesBytes,
            ),
            (
                ReadLimits::builder()
                    .max_relationship_xml_bytes(expected_relationships.bytes().len() - 1)
                    .unwrap()
                    .build()
                    .unwrap(),
                crate::ReadResource::RelationshipXmlBytes,
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
                Err(OpcError::ReadLimit { resource: actual, .. }) if actual == resource
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

        let after = package.source_relationships(&owner).unwrap();
        assert!(!after.member_present());
        assert_eq!(after.bytes(), replacement_relationships.bytes());

        let output = crate::PackageWriter::to_bytes(&package).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&output).unwrap();
        let relationship_member = owner.rels_uri().unwrap().membername().to_owned();
        assert!(!archive.file_names().any(|name| name == relationship_member));
        let reopened = OpcPackage::from_bytes(&output).unwrap();
        assert!(
            !reopened
                .source_relationships(&owner)
                .unwrap()
                .member_present()
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
        let pkg = OpcPackage::from_vec(create_minimal_docx()).unwrap();
        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn owned_constructors_preserve_mixed_storage_and_exact_source() {
        let source = with_eocd_comment(create_mixed_minimal_docx(), b"mixed storage");
        let file = NamedTempFile::new().unwrap();
        fs::write(file.path(), &source).unwrap();

        let from_vec = OpcPackage::from_vec(source.clone()).unwrap();
        let from_reader = OpcPackage::from_reader(Cursor::new(source.clone())).unwrap();
        let from_path = OpcPackage::open(file.path()).unwrap();
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
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();

        for result in [
            OpcPackage::from_vec_with_limits(source.clone(), limits),
            OpcPackage::from_reader_with_limits(Cursor::new(source.clone()), limits),
            OpcPackage::open_with_limits(file.path(), limits),
        ] {
            assert!(matches!(
                result,
                Err(OpcError::ReadLimit {
                    resource: crate::ReadResource::InputBytes,
                    actual: 4,
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
            OpcPackage::from_vec(source.clone()),
            OpcPackage::from_reader(Cursor::new(source.clone())),
            OpcPackage::open(file.path()),
        ] {
            assert!(matches!(result, Err(OpcError::ZipError(_))));
        }
    }

    #[test]
    fn clone_shares_owned_source_but_revocation_is_independent() {
        let bytes = with_eocd_comment(create_minimal_docx(), b"exact source");
        let package = OpcPackage::from_vec(bytes).expect("open owned package");
        let mut clone = package.clone();

        assert!(package.is_unmodified_owned_source());
        assert!(clone.is_unmodified_owned_source());
        assert!(!OpcPackage::new().is_unmodified_owned_source());
        assert!(
            !OpcPackage::from_bytes(&create_minimal_docx())
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
        let donor = OpcPackage::from_vec(bytes.clone()).unwrap();
        let shared = Arc::new(bytes.clone());
        let packages = [
            OpcPackage::from_vec(bytes.clone()).unwrap(),
            OpcPackage::from_reader(Cursor::new(bytes.clone())).unwrap(),
            OpcPackage::open(file.path()).unwrap(),
            OpcPackage::from_vec_reusing_payloads(bytes.clone(), ReadLimits::default(), &donor)
                .unwrap(),
            OpcPackage::from_shared_vec_reusing_payloads(
                Arc::clone(&shared),
                ReadLimits::default(),
                &donor,
            )
            .unwrap(),
            OpcPackage::from_shared_vec_with_limits(Arc::clone(&shared), ReadLimits::default())
                .unwrap(),
        ];
        for package in &packages {
            assert_eq!(package.exact_source_sha256(), Some(expected));
            assert_eq!(package.exact_source_len(), Some(bytes.len()));
            let mut streamed = Vec::new();
            package.to_stream(&mut streamed).unwrap();
            assert_eq!(sha256_of(&streamed), expected);
            assert_eq!(streamed, bytes);
        }

        assert_eq!(OpcPackage::new().exact_source_sha256(), None);
        assert_eq!(OpcPackage::new().exact_source_len(), None);
        assert_eq!(
            OpcPackage::from_bytes(&bytes)
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
        let package = OpcPackage::from_vec(bytes.clone()).unwrap();
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

        let again = OpcPackage::from_vec(bytes.clone()).unwrap();
        let again_source = again.source_archive.as_ref().unwrap();
        assert!(!again_source.shares_memo_with(source));
        assert!(!again_source.sha256_is_memoized());
        assert_eq!(again.exact_source_sha256(), package.exact_source_sha256());

        let other = with_eocd_comment(create_minimal_docx(), b"other bytes");
        let different = OpcPackage::from_vec(other.clone()).unwrap();
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
        let donor = OpcPackage::from_vec(bytes.clone()).unwrap();
        let shared = Arc::new(bytes.clone());

        let owned =
            OpcPackage::from_vec_reusing_payloads(bytes.clone(), ReadLimits::default(), &donor)
                .unwrap();
        let reused = OpcPackage::from_shared_vec_reusing_payloads(
            Arc::clone(&shared),
            ReadLimits::default(),
            &donor,
        )
        .unwrap();
        let deferred =
            OpcPackage::from_shared_vec_with_limits(Arc::clone(&shared), ReadLimits::default())
                .unwrap();
        for package in [&reused, &deferred] {
            assert!(package.is_unmodified_owned_source());
            assert!(Arc::ptr_eq(
                &package.exact_source_shared().unwrap(),
                &shared
            ));
            assert_eq!(package.source_read_limits(), Some(ReadLimits::default()));
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

        let tight = ReadLimits::builder()
            .max_input_bytes(8)
            .unwrap()
            .build()
            .unwrap();
        let handles = Arc::strong_count(&shared);
        let owned_refusal = OpcPackage::from_vec_with_limits(bytes.clone(), tight)
            .err()
            .map(|error| error.to_string());
        let shared_refusal = OpcPackage::from_shared_vec_with_limits(Arc::clone(&shared), tight)
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
    fn mutable_api_entries_revoke_owned_source_even_on_failure_or_noop() {
        let source = with_eocd_comment(create_minimal_docx(), b"exact source");
        let base = OpcPackage::from_vec(source).expect("open owned package");
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
    fn relationship_and_option_mutable_apis_revoke_owned_source() {
        let source = with_eocd_comment(create_minimal_docx(), b"exact source");
        let base = OpcPackage::from_vec(source).expect("open owned package");

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
        let mut package = OpcPackage::from_bytes(&create_minimal_docx()).unwrap();
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
        let original = OpcPackage::from_vec(bytes.clone()).unwrap();
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

        let mut replacement = OpcPackage::from_vec(bytes.clone()).unwrap();
        replacement.unsign();
        assert!(replacement.validate_signature_edit_from(&original).is_err());
        assert!(
            OpcPackage::new()
                .validate_signature_edit_from(&original)
                .is_err()
        );

        let mut borrowed = OpcPackage::from_bytes(&bytes).unwrap();
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
        let limits = ReadLimits::builder()
            .max_input_bytes(3)
            .unwrap()
            .build()
            .unwrap();
        assert!(matches!(
            OpcPackage::from_bytes_with_limits(b"four", limits),
            Err(OpcError::ReadLimit {
                resource: crate::ReadResource::InputBytes,
                actual: 4,
                maximum: 3,
            })
        ));
    }

    #[test]
    fn package_retains_ingress_limits_through_clone_and_edits() {
        let bytes = create_minimal_docx();
        let limits = ReadLimits::builder()
            .max_input_bytes(bytes.len() as u64)
            .unwrap()
            .max_part_bytes(8192)
            .unwrap()
            .build()
            .unwrap();
        let borrowed = OpcPackage::from_bytes_with_limits(&bytes, limits).unwrap();
        let owned = OpcPackage::from_vec_with_limits(bytes.clone(), limits).unwrap();
        let streamed = OpcPackage::from_reader_with_limits(Cursor::new(bytes), limits).unwrap();
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
        assert_eq!(OpcPackage::new().read_limits(), ReadLimits::default());
    }

    #[test]
    fn test_main_document_part() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor).unwrap();

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
        let mut reopened = OpcPackage::from_bytes(&bytes).unwrap();
        assert!(reopened.holds_original_source_xml(reopened.get_part(&uri).unwrap()));

        // Equal bytes in a newly allocated payload are still authored from the
        // writer's point of view. A byte comparison would incorrectly retain
        // the source proof here; allocation identity must revoke it.
        let equal_copy = reopened.get_part(&uri).unwrap().blob().to_vec();
        reopened.get_part_mut(&uri).unwrap().set_blob(equal_copy);
        assert!(!reopened.holds_original_source_xml(reopened.get_part(&uri).unwrap()));
    }
}
