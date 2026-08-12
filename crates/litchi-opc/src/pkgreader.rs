//! Low-level, read-only API to a serialized Open Packaging Convention (OPC) package.
//!
//! This module provides the `PackageReader` for parsing OPC packages, including
//! content type mapping, relationship resolution, and part loading. It uses
//! efficient algorithms for parsing and minimal memory allocation.

use crate::constants::{content_type as ct, namespace};
use crate::content_type::ContentTypeMap;
use crate::error::{OpcError, Result};
use crate::execution::OpenSession;
use crate::limits::{ReadLimits, ReadResource};
use crate::members::{NonPartMember, NonPartReason, PartNameIndex, part_name_for_member};
use crate::packuri::{PACKAGE_URI, PackURI};
use crate::phys_pkg::PhysPkgReader;
use crate::rel::{TargetMode, relationship_target_components};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::collections::{HashMap, HashSet, TryReserveError};

/// The small ZIP surface needed by the structural OPC reader.
///
/// Keeping this behind a private trait lets the eager byte-slice ingress and
/// the source-backed ingress share exactly the same conformance checks.
pub(crate) trait ArchiveAccess {
    fn len(&self) -> usize;
    fn contains(&self, name: &str) -> bool;
    fn file_names(&self) -> Box<dyn Iterator<Item = &str> + '_>;
    fn metadata(
        &self,
        name: &str,
    ) -> std::result::Result<soapberry_zip::office::Metadata, soapberry_zip::Error>;
    fn read(&self, name: &str) -> std::result::Result<Vec<u8>, soapberry_zip::Error>;
}

impl ArchiveAccess for soapberry_zip::office::LazyArchiveReader<'_> {
    fn len(&self) -> usize {
        Self::len(self)
    }

    fn contains(&self, name: &str) -> bool {
        Self::contains(self, name)
    }

    fn file_names(&self) -> Box<dyn Iterator<Item = &str> + '_> {
        Box::new(Self::file_names(self))
    }

    fn metadata(
        &self,
        name: &str,
    ) -> std::result::Result<soapberry_zip::office::Metadata, soapberry_zip::Error> {
        Self::metadata(self, name)
    }

    fn read(&self, name: &str) -> std::result::Result<Vec<u8>, soapberry_zip::Error> {
        Self::read(self, name)
    }
}

impl<R: soapberry_zip::ReaderAt> ArchiveAccess for soapberry_zip::office::IndexedArchive<R> {
    fn len(&self) -> usize {
        Self::len(self)
    }

    fn contains(&self, name: &str) -> bool {
        Self::contains(self, name)
    }

    fn file_names(&self) -> Box<dyn Iterator<Item = &str> + '_> {
        Box::new(Self::file_names(self))
    }

    fn metadata(
        &self,
        name: &str,
    ) -> std::result::Result<soapberry_zip::office::Metadata, soapberry_zip::Error> {
        Self::metadata(self, name)
    }

    fn read(&self, name: &str) -> std::result::Result<Vec<u8>, soapberry_zip::Error> {
        Self::read(self, name)
    }
}

/// Reserved ZIP item name of the content types stream (ECMA-376 Part 2 §10.1.2.2).
const CONTENT_TYPES_MEMBER: &str = "[Content_Types].xml";

/// Serialized part with its content and relationships.
///
/// Represents a part as loaded from the physical package, before
/// being converted into a Part object.
#[derive(Debug)]
pub struct SerializedPart {
    /// The partname (URI) of this part
    pub partname: PackURI,

    /// The content type of this part
    pub content_type: String,

    /// The binary content of this part
    pub blob: Vec<u8>,

    /// Serialized relationships from this part
    /// Uses `SmallVec` for efficient storage of typically small relationship collections
    pub srels: SmallVec<[SerializedRelationship; 8]>,
}

/// Structural information retained by the source-backed reader for one part.
#[derive(Debug)]
pub(crate) struct DeferredPart {
    pub(crate) partname: PackURI,
    pub(crate) content_type: String,
    pub(crate) srels: SmallVec<[SerializedRelationship; 8]>,
}

/// Fully validated package catalog whose ordinary part payloads remain in ZIP.
#[derive(Debug)]
pub(crate) struct SourceCatalog {
    pub(crate) pkg_srels: SmallVec<[SerializedRelationship; 8]>,
    pub(crate) parts: Vec<DeferredPart>,
    pub(crate) non_part_members: Vec<NonPartMember>,
}

/// Exact source-catalog phase used only by the validation entry point.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum ValidationCatalogPhase {
    /// ZIP indexing and package-level admission.
    Ingress,
    /// `[Content_Types].xml` and physical part-catalog admission.
    Catalog,
    /// Relationship manifests loaded for the package and admitted typed parts.
    LoadedRelationships,
}

/// A source-catalog failure with validation-only phase provenance.
#[derive(Debug)]
pub(crate) struct ValidationCatalogError {
    pub(crate) phase: ValidationCatalogPhase,
    pub(crate) error: OpcError,
}

#[cfg(test)]
mod physical_part_tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::PackageReader;
    use crate::{BlobPart, OpcPackage, PackURI, PackageWriter};

    const ORPHAN_PART_NAME: &str = "/custom/orphan.xml";
    const ORPHAN_CONTENT: &[u8] = b"<extension xmlns=\"urn:litchi:test\">preserve me</extension>";

    #[test]
    fn recognizes_only_reserved_relationship_members() {
        assert!(PackageReader::is_relationship_member("_rels/.rels"));
        assert!(PackageReader::is_relationship_member(
            "word/_rels/document.xml.rels"
        ));
        assert!(!PackageReader::is_relationship_member("custom/data.rels"));
        assert!(!PackageReader::is_relationship_member("_rels/data.xml"));
    }

    #[test]
    fn preserves_unreferenced_physical_part_across_round_trip() {
        let mut source = OpcPackage::new();
        source.add_part(Box::new(BlobPart::new(
            PackURI::new(ORPHAN_PART_NAME).unwrap(),
            "application/vnd.litchi.extension+xml".to_string(),
            ORPHAN_CONTENT.to_vec(),
        )));

        let serialized = PackageWriter::to_bytes(&source).unwrap();
        let loaded = OpcPackage::from_bytes(&serialized).unwrap();
        let orphan = loaded
            .iter_parts()
            .find(|part| part.partname().as_str() == ORPHAN_PART_NAME)
            .expect("unreferenced physical part must be loaded");
        assert_eq!(orphan.blob(), ORPHAN_CONTENT);

        let reserialized = PackageWriter::to_bytes(&loaded).unwrap();
        let reloaded = OpcPackage::from_bytes(&reserialized).unwrap();
        let surviving_orphan = reloaded
            .iter_parts()
            .find(|part| part.partname().as_str() == ORPHAN_PART_NAME)
            .expect("unreferenced physical part must survive save and reopen");
        assert_eq!(surviving_orphan.blob(), ORPHAN_CONTENT);
    }
}

/// Serialized relationship as read from a .rels file.
///
/// Contains all relationship information in string form, before
/// being converted into Relationship objects with resolved part references.
#[derive(Debug, Clone)]
pub struct SerializedRelationship {
    /// Base URI for resolving relative references
    pub base_uri: String,

    /// Full source part URI, when known.
    pub source_uri: Option<String>,

    /// Relationship ID (e.g., "rId1")
    pub r_id: String,

    /// Relationship type URI
    pub reltype: String,

    /// Target reference (relative URI or external URL)
    pub target_ref: String,

    /// Target mode (Internal or External)
    pub target_mode: TargetMode,
}

impl SerializedRelationship {
    /// Check if this is an external relationship.
    #[inline]
    #[must_use]
    pub fn is_external(&self) -> bool {
        self.target_mode == TargetMode::External
    }

    /// Get the target partname for internal relationships.
    ///
    /// Resolves the relative target reference against the base URI
    /// to produce an absolute `PackURI`.
    ///
    /// # Errors
    /// Returns an error for external relationships, for internal targets with no
    /// resolvable part path, or when the resolved name is not a valid `PackURI`.
    pub fn target_partname(&self) -> Result<PackURI> {
        if self.is_external() {
            return Err(OpcError::InvalidRelationship(
                "Cannot get target_partname for external relationship".to_string(),
            ));
        }
        let path = relationship_target_components(&self.target_ref).0;
        if path.is_empty() {
            return self
                .source_uri
                .as_deref()
                .filter(|source| *source != "/")
                .ok_or_else(|| {
                    OpcError::InvalidRelationship(
                        "Internal relationship target has no part path".to_string(),
                    )
                })
                .and_then(|source| PackURI::new(source).map_err(OpcError::InvalidPackUri));
        }
        PackURI::from_rel_ref(&self.base_uri, path).map_err(OpcError::InvalidPackUri)
    }
}

/// Package reader that provides access to serialized parts and relationships.
///
/// This is the main entry point for reading OPC packages. It handles parsing
/// the package structure, resolving relationships, and loading parts efficiently.
pub struct PackageReader {
    /// Package-level relationships
    /// Uses `SmallVec` for efficient storage of typically small relationship collections
    pkg_srels: SmallVec<[SerializedRelationship; 8]>,

    /// All serialized parts in the package
    sparts: Vec<SerializedPart>,

    /// ZIP items that were present but are not OPC parts
    non_part_members: Vec<NonPartMember>,
}

impl PackageReader {
    /// Open and parse an OPC package from a byte slice.
    ///
    /// Uses lazy decompression for maximum throughput:
    /// 1. Decompress files on-demand during relationship graph traversal
    /// 2. Parse each file as soon as it's decompressed (pipelining)
    /// 3. Cache decompressed data to avoid redundant work
    ///
    /// This approach is faster than pre-loading everything because:
    /// - Parsing can start while other files are still being decompressed
    /// - Files not in the relationship graph are never decompressed
    /// - Memory pressure is reduced (don't hold all decompressed data at once)
    ///
    /// # Arguments
    /// * `phys_reader` - Physical package reader for accessing ZIP contents
    ///
    /// # Returns
    /// A new `PackageReader` with all parts and relationships loaded
    ///
    /// # Errors
    /// Returns an error when the package violates the OPC specification (missing
    /// content types stream, malformed relationships manifest, invalid part
    /// names or content types), when a ZIP member cannot be read, or when any
    /// configured read limit is exceeded.
    pub fn from_phys_reader(phys_reader: &PhysPkgReader<'_>) -> Result<Self> {
        let archive = phys_reader.archive();
        let limits = phys_reader.limits();
        limits.check(
            ReadResource::ArchiveMembers,
            archive.len() as u64,
            limits.max_archive_members() as u64,
        )?;

        let relationship_part_count = archive
            .file_names()
            .filter(|member_name| Self::is_relationship_member(member_name))
            .count();
        limits.check(
            ReadResource::RelationshipParts,
            relationship_part_count as u64,
            limits.max_relationship_parts() as u64,
        )?;
        let mut relationship_ledger = RelationshipLedger::default();

        // Phase 1: Decompress and parse content types (on-demand)
        let content_types_member = Self::locate_content_types_member(archive)?;
        let content_types_metadata = archive.metadata(content_types_member)?;
        limits.check(
            ReadResource::ContentTypesBytes,
            content_types_metadata.uncompressed_size(),
            limits.max_content_types_bytes() as u64,
        )?;
        let content_types_xml = archive.read(content_types_member)?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml, limits)?;

        // Phase 2: Get package-level relationships (on-demand decompression)
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels =
            Self::load_rels_lazy(archive, &package_uri, limits, &mut relationship_ledger)?;

        // Phase 3: Load every physical part. Relationship traversal alone is
        // insufficient because OPC permits parts with no incoming relationship.
        let mut non_part_members = Vec::new();
        non_part_members
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC non-part members", source))?;
        let sparts = Self::load_parts_lazy(
            archive,
            content_types_member,
            &pkg_srels,
            &content_types,
            &mut non_part_members,
            limits,
            &mut relationship_ledger,
            |names| {
                Ok(names
                    .iter()
                    .map(|name| (*name, archive.read(name)))
                    .collect())
            },
        )?;

        Ok(Self {
            pkg_srels,
            sparts,
            non_part_members,
        })
    }

    /// Open a physical package with an explicitly scheduled eager bulk-read session.
    ///
    /// Ordinary constructors call [`Self::from_phys_reader`] and remain serial.
    pub(crate) fn from_phys_reader_with_session(
        phys_reader: &PhysPkgReader<'_>,
        session: &OpenSession,
    ) -> Result<Self> {
        let archive = phys_reader.archive();
        let limits = phys_reader.limits();
        limits.check(
            ReadResource::ArchiveMembers,
            archive.len() as u64,
            limits.max_archive_members() as u64,
        )?;

        let relationship_part_count = archive
            .file_names()
            .filter(|member_name| Self::is_relationship_member(member_name))
            .count();
        limits.check(
            ReadResource::RelationshipParts,
            relationship_part_count as u64,
            limits.max_relationship_parts() as u64,
        )?;
        let mut relationship_ledger = RelationshipLedger::default();

        let content_types_member = Self::locate_content_types_member(archive)?;
        let content_types_metadata = archive.metadata(content_types_member)?;
        limits.check(
            ReadResource::ContentTypesBytes,
            content_types_metadata.uncompressed_size(),
            limits.max_content_types_bytes() as u64,
        )?;
        let content_types_xml = archive.read(content_types_member)?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml, limits)?;

        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels =
            Self::load_rels_lazy(archive, &package_uri, limits, &mut relationship_ledger)?;

        let mut non_part_members = Vec::new();
        non_part_members
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC non-part members", source))?;
        let sparts = Self::load_parts_lazy(
            archive,
            content_types_member,
            &pkg_srels,
            &content_types,
            &mut non_part_members,
            limits,
            &mut relationship_ledger,
            |names| session.read_many(archive, names),
        )?;

        Ok(Self {
            pkg_srels,
            sparts,
            non_part_members,
        })
    }

    /// Resolve the ZIP item that holds the content types stream.
    ///
    /// The reserved name is `[Content_Types].xml`. ECMA-376 Part 2 §9.1.1.2 makes
    /// item-name comparison ASCII case-insensitive, so a package that stores the
    /// stream as `[content_types].xml` is still unambiguous; Apache POI resolves
    /// it the same way. The exact spelling wins when both are present.
    fn locate_content_types_member<A: ArchiveAccess + ?Sized>(archive: &A) -> Result<&str> {
        if archive.contains(CONTENT_TYPES_MEMBER) {
            return Ok(CONTENT_TYPES_MEMBER);
        }
        archive
            .file_names()
            .find(|name| name.eq_ignore_ascii_case(CONTENT_TYPES_MEMBER))
            .ok_or_else(|| OpcError::PartNotFound(CONTENT_TYPES_MEMBER.to_string()))
    }

    /// Parse relationships XML into `SerializedRelationship` structs.
    #[cfg(test)]
    fn parse_rels_xml(
        rels_xml: &[u8],
        base_uri: &str,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let limits = ReadLimits::default();
        let mut ledger = RelationshipLedger::default();
        ledger.preflight_xml_bytes(limits, rels_xml.len() as u64)?;
        ledger.retain_xml_bytes(limits, rels_xml.len() as u64)?;
        Self::parse_rels_xml_with_source(rels_xml, base_uri, None, limits, &mut ledger)
    }

    fn parse_rels_xml_with_source(
        rels_xml: &[u8],
        base_uri: &str,
        source_uri: Option<&str>,
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let mut srels = SmallVec::new();
        let mut reader = NsReader::from_reader(rels_xml);
        reader.config_mut().trim_text(true);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut root_seen = false;
        let mut ids = HashSet::new();
        let mut events = 0usize;

        loop {
            events = checked_increment(events, limits.max_xml_events(), ReadResource::XmlEvents)?;
            ledger.add_event(limits)?;
            let decoder = reader.decoder();
            let (resolved_namespace, event) = reader.read_resolved_event()?;
            match event {
                Event::Start(element) => {
                    let next_depth =
                        checked_increment(depth, limits.max_xml_depth(), ReadResource::XmlDepth)?;
                    inspect_relationship_element(
                        &mut srels,
                        &mut ids,
                        &resolved_namespace,
                        &element,
                        decoder,
                        base_uri,
                        source_uri,
                        depth,
                        &mut root_seen,
                        limits,
                        ledger,
                    )?;
                    depth = next_depth;
                },
                Event::Empty(element) => inspect_relationship_element(
                    &mut srels,
                    &mut ids,
                    &resolved_namespace,
                    &element,
                    decoder,
                    base_uri,
                    source_uri,
                    depth,
                    &mut root_seen,
                    limits,
                    ledger,
                )?,
                Event::End(_) => {
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        OpcError::InvalidRelationshipsManifest(
                            "unmatched closing element".to_string(),
                        )
                    })?;
                },
                Event::Text(text) if depth <= 1 && !text.as_ref().is_empty() => {
                    return Err(OpcError::InvalidRelationshipsManifest(
                        "text is not permitted under Relationships".to_string(),
                    ));
                },
                Event::DocType(_) => {
                    return Err(OpcError::InvalidRelationshipsManifest(
                        "DTDs are not permitted in relationships parts".to_string(),
                    ));
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::GeneralRef(_) => {},
            }
        }

        if !root_seen || depth != 0 {
            return Err(OpcError::InvalidRelationshipsManifest(
                "missing or unclosed Relationships root".to_string(),
            ));
        }
        if base_uri == "/"
            && srels
                .iter()
                .filter(|relationship| is_core_properties_relationship(&relationship.reltype))
                .count()
                > 1
        {
            return Err(OpcError::MultipleCorePropertiesRelationships);
        }
        Ok(srels)
    }

    /// Load relationships using lazy on-demand decompression.
    ///
    /// Decompresses and parses the relationships file for a given source URI.
    /// The result is cached by the lazy archive reader for subsequent access.
    fn load_rels_lazy<A: ArchiveAccess + ?Sized>(
        archive: &A,
        source_uri: &PackURI,
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let rels_uri = source_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let rels_path = rels_uri.membername();

        match archive.metadata(rels_path) {
            Ok(metadata) => {
                ledger.preflight_xml_bytes(limits, metadata.uncompressed_size())?;
                let rels_xml = archive.read(rels_path)?;
                ledger.retain_xml_bytes(limits, rels_xml.len() as u64)?;
                Self::parse_rels_xml_with_source(
                    &rels_xml,
                    source_uri.base_uri(),
                    Some(source_uri.as_str()),
                    limits,
                    ledger,
                )
            },
            Err(error) if is_member_missing(&error) => Ok(SmallVec::new()),
            Err(error) => Err(error.into()),
        }
    }

    /// Load all physical parts using parallel decompression.
    ///
    /// The ZIP members are the authority on which parts exist (ECMA-376 Part 2
    /// §10.1.3): relationships are edges recorded on their source part, not
    /// evidence that a target part is present. The phases are:
    /// 1. Walk the relationship graph to validate targets and learn which part
    ///    names the package refers to.
    /// 2. Classify every ZIP member into a part or a reported non-part member.
    /// 3. Check the resulting part-name collection for OPC name conflicts.
    /// 4. Decompress all part contents in parallel.
    fn load_parts_lazy<A, ReadMany>(
        archive: &A,
        content_types_member: &str,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
        non_part_members: &mut Vec<NonPartMember>,
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
        read_many: ReadMany,
    ) -> Result<Vec<SerializedPart>>
    where
        A: ArchiveAccess + ?Sized,
        ReadMany: for<'name> FnOnce(
            &'name [&'name str],
        ) -> Result<
            Vec<(
                &'name str,
                std::result::Result<Vec<u8>, soapberry_zip::Error>,
            )>,
        >,
    {
        // Phase 1: relationship reachability. Relationship types belong to
        // edges, not parts, so they are intentionally not recorded here.
        let mut relationships = Self::walk_relationship_graph(archive, pkg_srels, limits, ledger)?;

        // Phase 2 and 3: classify members, then admit the survivors as parts.
        let mut index = PartNameIndex::try_with_capacity(archive.len())?;
        let mut typed_parts: Vec<(PackURI, String)> = Vec::new();
        typed_parts
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC typed parts", source))?;
        for member_name in archive.file_names() {
            if member_name.is_empty()
                || member_name.ends_with('/')
                || member_name == content_types_member
            {
                continue;
            }

            if Self::relationship_part_has_relationships(member_name) {
                return Err(OpcError::RelationshipPartCannotBeSource(
                    member_name.to_string(),
                ));
            }

            let max_member_name_bytes =
                usize::try_from(limits.max_archive_member_name_bytes()).unwrap_or(usize::MAX);
            let Some(partname) = part_name_for_member(member_name, max_member_name_bytes) else {
                non_part_members.push(NonPartMember::new(
                    member_name,
                    NonPartReason::UnmappablePartName,
                )?);
                continue;
            };

            let is_relationship_part = Self::is_relationship_member(partname.membername());
            let content_type = if is_relationship_part {
                Self::relationship_part_content_type(content_types, &partname)?
            } else {
                match content_types.get(&partname) {
                    Ok(content_type) => content_type,
                    // An untyped item that nothing refers to is archive junk,
                    // not a non-conforming part (ECMA-376 Part 2 §10.1.2.2).
                    Err(OpcError::ContentTypeNotFound(_))
                        if !relationships.contains_key(partname.as_str()) =>
                    {
                        non_part_members.push(NonPartMember::new(
                            member_name,
                            NonPartReason::UntypedAndUnreferenced,
                        )?);
                        continue;
                    },
                    Err(error) => return Err(error),
                }
            };

            // Relationship parts are named parts for conflict purposes even
            // though they are surfaced through their source part, not on
            // their own.
            index.insert(&partname)?;
            if !is_relationship_part {
                let part_count =
                    checked_increment(typed_parts.len(), limits.max_parts(), ReadResource::Parts)?;
                debug_assert_eq!(part_count, typed_parts.len() + 1);
                typed_parts.push((partname, content_type));
            }
        }

        // Phase 4: parallel decompression of every admitted part.
        let mut member_names = Vec::new();
        member_names
            .try_reserve_exact(typed_parts.len())
            .map_err(|source| allocation("OPC member-name batch", source))?;
        member_names.extend(
            typed_parts
                .iter()
                .map(|(partname, _)| partname.membername()),
        );
        let mut declared_part_bytes = 0u64;
        for member_name in &member_names {
            let declared = archive.metadata(member_name)?.uncompressed_size();
            limits.check(ReadResource::PartBytes, declared, limits.max_part_bytes())?;
            declared_part_bytes = checked_add(
                declared_part_bytes,
                declared,
                ReadResource::TotalPartBytes,
                limits.max_total_part_bytes(),
            )?;
        }
        let mut decompressed = HashMap::new();
        decompressed
            .try_reserve(typed_parts.len())
            .map_err(|source| allocation("OPC decompressed parts", source))?;
        let mut retained_part_bytes = 0u64;
        for (member_name, result) in read_many(&member_names)? {
            let blob = result?;
            limits.check(
                ReadResource::PartBytes,
                blob.len() as u64,
                limits.max_part_bytes(),
            )?;
            retained_part_bytes = checked_add(
                retained_part_bytes,
                blob.len() as u64,
                ReadResource::TotalPartBytes,
                limits.max_total_part_bytes(),
            )?;
            decompressed.insert(member_name.to_string(), blob);
        }

        // Phase 5: build SerializedPart structures (take ownership, no cloning)
        let mut sparts = Vec::new();
        sparts
            .try_reserve_exact(typed_parts.len())
            .map_err(|source| allocation("OPC serialized parts", source))?;
        for (partname, content_type) in typed_parts {
            let srels = match relationships.remove(partname.as_str()) {
                Some(srels) => srels,
                None => Self::load_rels_lazy(archive, &partname, limits, ledger)?,
            };
            // Remove from map to take ownership instead of cloning
            let blob = decompressed
                .remove(partname.membername())
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
            sparts.push(SerializedPart {
                partname,
                content_type,
                blob,
                srels,
            });
        }

        Ok(sparts)
    }

    /// Perform the same structural admission as [`Self::load_parts_lazy`]
    /// without reading ordinary part payloads.  This is deliberately kept next
    /// to the eager path so content-type, relationship, classification, and
    /// name-conflict semantics cannot drift between the two ingress modes.
    fn load_part_catalog<A: ArchiveAccess + ?Sized>(
        archive: &A,
        content_types_member: &str,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
        non_part_members: &mut Vec<NonPartMember>,
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
    ) -> Result<Vec<DeferredPart>> {
        let mut relationships = Self::walk_relationship_graph(archive, pkg_srels, limits, ledger)?;
        let mut index = PartNameIndex::try_with_capacity(archive.len())?;
        let mut typed_parts: Vec<(PackURI, String)> = Vec::new();
        typed_parts
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC deferred typed parts", source))?;

        let mut declared_part_bytes = 0u64;
        for member_name in archive.file_names() {
            if member_name.is_empty()
                || member_name.ends_with('/')
                || member_name == content_types_member
            {
                continue;
            }
            if Self::relationship_part_has_relationships(member_name) {
                return Err(OpcError::RelationshipPartCannotBeSource(
                    member_name.to_string(),
                ));
            }
            let max_member_name_bytes =
                usize::try_from(limits.max_archive_member_name_bytes()).unwrap_or(usize::MAX);
            let Some(partname) = part_name_for_member(member_name, max_member_name_bytes) else {
                non_part_members.push(NonPartMember::new(
                    member_name,
                    NonPartReason::UnmappablePartName,
                )?);
                continue;
            };
            let is_relationship_part = Self::is_relationship_member(partname.membername());
            let content_type = if is_relationship_part {
                Self::relationship_part_content_type(content_types, &partname)?
            } else {
                match content_types.get(&partname) {
                    Ok(content_type) => content_type,
                    Err(OpcError::ContentTypeNotFound(_))
                        if !relationships.contains_key(partname.as_str()) =>
                    {
                        non_part_members.push(NonPartMember::new(
                            member_name,
                            NonPartReason::UntypedAndUnreferenced,
                        )?);
                        continue;
                    },
                    Err(error) => return Err(error),
                }
            };
            index.insert(&partname)?;
            if !is_relationship_part {
                let part_count =
                    checked_increment(typed_parts.len(), limits.max_parts(), ReadResource::Parts)?;
                debug_assert_eq!(part_count, typed_parts.len() + 1);
                let declared = archive.metadata(partname.membername())?.uncompressed_size();
                limits.check(ReadResource::PartBytes, declared, limits.max_part_bytes())?;
                declared_part_bytes = checked_add(
                    declared_part_bytes,
                    declared,
                    ReadResource::TotalPartBytes,
                    limits.max_total_part_bytes(),
                )?;
                typed_parts.push((partname, content_type));
            }
        }

        let mut parts = Vec::new();
        parts
            .try_reserve_exact(typed_parts.len())
            .map_err(|source| allocation("OPC deferred parts", source))?;
        for (partname, content_type) in typed_parts {
            let srels = match relationships.remove(partname.as_str()) {
                Some(srels) => srels,
                None => Self::load_rels_lazy(archive, &partname, limits, ledger)?,
            };
            parts.push(DeferredPart {
                partname,
                content_type,
                srels,
            });
        }
        Ok(parts)
    }

    fn load_part_catalog_for_validation<A: ArchiveAccess + ?Sized>(
        archive: &A,
        content_types_member: &str,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
        non_part_members: &mut Vec<NonPartMember>,
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
    ) -> std::result::Result<Vec<DeferredPart>, ValidationCatalogError> {
        let phase = |phase, error| ValidationCatalogError { phase, error };
        let mut relationships =
            Self::walk_relationship_graph(archive, pkg_srels, limits, ledger)
                .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
        let mut index = PartNameIndex::try_with_capacity(archive.len())
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let mut typed_parts: Vec<(PackURI, String)> = Vec::new();
        typed_parts
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC deferred typed parts", source))
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;

        let mut declared_part_bytes = 0u64;
        for member_name in archive.file_names() {
            if member_name.is_empty()
                || member_name.ends_with('/')
                || member_name == content_types_member
            {
                continue;
            }
            if Self::relationship_part_has_relationships(member_name) {
                return Err(phase(
                    ValidationCatalogPhase::Catalog,
                    OpcError::RelationshipPartCannotBeSource(member_name.to_string()),
                ));
            }
            let max_member_name_bytes =
                usize::try_from(limits.max_archive_member_name_bytes()).unwrap_or(usize::MAX);
            let Some(partname) = part_name_for_member(member_name, max_member_name_bytes) else {
                non_part_members.push(
                    NonPartMember::new(member_name, NonPartReason::UnmappablePartName)
                        .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?,
                );
                continue;
            };
            let is_relationship_part = Self::is_relationship_member(partname.membername());
            let content_type = if is_relationship_part {
                Self::relationship_part_content_type(content_types, &partname)
                    .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?
            } else {
                match content_types.get(&partname) {
                    Ok(content_type) => content_type,
                    Err(OpcError::ContentTypeNotFound(_))
                        if !relationships.contains_key(partname.as_str()) =>
                    {
                        non_part_members.push(
                            NonPartMember::new(member_name, NonPartReason::UntypedAndUnreferenced)
                                .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?,
                        );
                        continue;
                    },
                    Err(error) => return Err(phase(ValidationCatalogPhase::Catalog, error)),
                }
            };
            index
                .insert(&partname)
                .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
            if !is_relationship_part {
                let part_count =
                    checked_increment(typed_parts.len(), limits.max_parts(), ReadResource::Parts)
                        .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
                debug_assert_eq!(part_count, typed_parts.len() + 1);
                let declared = archive
                    .metadata(partname.membername())
                    .map_err(OpcError::from)
                    .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?
                    .uncompressed_size();
                limits
                    .check(ReadResource::PartBytes, declared, limits.max_part_bytes())
                    .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
                declared_part_bytes = checked_add(
                    declared_part_bytes,
                    declared,
                    ReadResource::TotalPartBytes,
                    limits.max_total_part_bytes(),
                )
                .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
                typed_parts.push((partname, content_type));
            }
        }

        let mut parts = Vec::new();
        parts
            .try_reserve_exact(typed_parts.len())
            .map_err(|source| allocation("OPC deferred parts", source))
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        for (partname, content_type) in typed_parts {
            let srels = match relationships.remove(partname.as_str()) {
                Some(srels) => srels,
                None => Self::load_rels_lazy(archive, &partname, limits, ledger)
                    .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?,
            };
            parts.push(DeferredPart {
                partname,
                content_type,
                srels,
            });
        }
        Ok(parts)
    }

    pub(crate) fn source_catalog<A: ArchiveAccess + ?Sized>(
        archive: &A,
        limits: ReadLimits,
    ) -> Result<SourceCatalog> {
        limits.check(
            ReadResource::ArchiveMembers,
            archive.len() as u64,
            limits.max_archive_members() as u64,
        )?;
        let relationship_part_count = archive
            .file_names()
            .filter(|member_name| Self::is_relationship_member(member_name))
            .count();
        limits.check(
            ReadResource::RelationshipParts,
            relationship_part_count as u64,
            limits.max_relationship_parts() as u64,
        )?;
        let mut ledger = RelationshipLedger::default();
        let content_types_member = Self::locate_content_types_member(archive)?;
        let content_types_metadata = archive.metadata(content_types_member)?;
        limits.check(
            ReadResource::ContentTypesBytes,
            content_types_metadata.uncompressed_size(),
            limits.max_content_types_bytes() as u64,
        )?;
        let content_types_xml = archive.read(content_types_member)?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml, limits)?;
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels = Self::load_rels_lazy(archive, &package_uri, limits, &mut ledger)?;
        let mut non_part_members = Vec::new();
        non_part_members
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC non-part members", source))?;
        let parts = Self::load_part_catalog(
            archive,
            content_types_member,
            &pkg_srels,
            &content_types,
            &mut non_part_members,
            limits,
            &mut ledger,
        )?;
        Ok(SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
        })
    }

    /// Source-backed catalog admission with exact validation-only phase
    /// provenance. Ordinary package opens continue to use [`Self::source_catalog`].
    pub(crate) fn source_catalog_for_validation<A: ArchiveAccess + ?Sized>(
        archive: &A,
        limits: ReadLimits,
    ) -> std::result::Result<SourceCatalog, ValidationCatalogError> {
        let phase = |phase, error| ValidationCatalogError { phase, error };
        limits
            .check(
                ReadResource::ArchiveMembers,
                archive.len() as u64,
                limits.max_archive_members() as u64,
            )
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let relationship_part_count = archive
            .file_names()
            .filter(|member_name| Self::is_relationship_member(member_name))
            .count();
        limits
            .check(
                ReadResource::RelationshipParts,
                relationship_part_count as u64,
                limits.max_relationship_parts() as u64,
            )
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let mut ledger = RelationshipLedger::default();
        let content_types_member = Self::locate_content_types_member(archive)
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let content_types_metadata = archive
            .metadata(content_types_member)
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        limits
            .check(
                ReadResource::ContentTypesBytes,
                content_types_metadata.uncompressed_size(),
                limits.max_content_types_bytes() as u64,
            )
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let content_types_xml = archive
            .read(content_types_member)
            .map_err(OpcError::from)
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml, limits)
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let package_uri = PackURI::new(PACKAGE_URI)
            .map_err(OpcError::InvalidPackUri)
            .map_err(|error| phase(ValidationCatalogPhase::Ingress, error))?;
        let pkg_srels = Self::load_rels_lazy(archive, &package_uri, limits, &mut ledger)
            .map_err(|error| phase(ValidationCatalogPhase::LoadedRelationships, error))?;
        let mut non_part_members = Vec::new();
        non_part_members
            .try_reserve(archive.len())
            .map_err(|source| allocation("OPC non-part members", source))
            .map_err(|error| phase(ValidationCatalogPhase::Catalog, error))?;
        let parts = Self::load_part_catalog_for_validation(
            archive,
            content_types_member,
            &pkg_srels,
            &content_types,
            &mut non_part_members,
            limits,
            &mut ledger,
        )?;
        Ok(SourceCatalog {
            pkg_srels,
            parts,
            non_part_members,
        })
    }

    /// Walk the relationship graph, returning each visited part name with its
    /// own relationships.
    ///
    /// Only small `.rels` members are decompressed. A target that has no ZIP
    /// member is still visited: OPC defines no rule requiring a relationship
    /// target to resolve, and the dangling relationship stays visible on its
    /// source part.
    fn walk_relationship_graph<A: ArchiveAccess + ?Sized>(
        archive: &A,
        pkg_srels: &[SerializedRelationship],
        limits: ReadLimits,
        ledger: &mut RelationshipLedger,
    ) -> Result<HashMap<String, SmallVec<[SerializedRelationship; 8]>>> {
        let mut visited: HashMap<String, SmallVec<[SerializedRelationship; 8]>> = HashMap::new();
        visited
            .try_reserve(pkg_srels.len())
            .map_err(|source| allocation("OPC relationship graph", source))?;
        let mut work_queue: Vec<PackURI> = Vec::new();
        work_queue
            .try_reserve(pkg_srels.len())
            .map_err(|source| allocation("OPC relationship work queue", source))?;
        for srel in pkg_srels {
            Self::enqueue_target(srel, &mut visited, &mut work_queue, limits)?;
        }

        while let Some(partname) = work_queue.pop() {
            let part_srels = Self::load_rels_lazy(archive, &partname, limits, ledger)?;
            for child_srel in &part_srels {
                Self::enqueue_target(child_srel, &mut visited, &mut work_queue, limits)?;
            }
            visited.insert(partname.to_string(), part_srels);
        }

        Ok(visited)
    }

    /// Queue an internal relationship target for traversal, once.
    fn enqueue_target(
        srel: &SerializedRelationship,
        visited: &mut HashMap<String, SmallVec<[SerializedRelationship; 8]>>,
        work_queue: &mut Vec<PackURI>,
        limits: ReadLimits,
    ) -> Result<()> {
        if srel.is_external() {
            return Ok(());
        }
        let partname = srel.target_partname()?;
        limits.check(
            ReadResource::RelationshipTargetBytes,
            partname.as_str().len() as u64,
            limits.max_relationship_target_bytes() as u64,
        )?;
        if visited.contains_key(partname.as_str()) {
            return Ok(());
        }
        checked_increment(
            visited.len(),
            limits.max_relationship_graph_nodes(),
            ReadResource::RelationshipGraphNodes,
        )?;
        visited
            .try_reserve(1)
            .map_err(|source| allocation("OPC relationship graph", source))?;
        work_queue
            .try_reserve(1)
            .map_err(|source| allocation("OPC relationship work queue", source))?;
        visited.insert(partname.to_string(), SmallVec::new());
        work_queue.push(partname);
        Ok(())
    }

    /// Resolve the content type of a relationship part.
    ///
    /// ECMA-376 Part 2 §9.2 fixes the content type of every Relationships part,
    /// and §9.1.2 reserves the `_rels/*.rels` naming that identifies one, so a
    /// manifest that omits the mapping leaves nothing ambiguous. A manifest that
    /// maps a reserved relationship name onto some *other* type contradicts the
    /// specification and is rejected.
    fn relationship_part_content_type(
        content_types: &ContentTypeMap,
        partname: &PackURI,
    ) -> Result<String> {
        match content_types.get(partname) {
            Ok(declared) if declared == ct::OPC_RELATIONSHIPS => Ok(declared),
            Ok(declared) => Err(OpcError::InvalidContentType {
                value: declared,
                reason: format!("relationship part '{partname}' must be typed {}", {
                    ct::OPC_RELATIONSHIPS
                }),
            }),
            Err(OpcError::ContentTypeNotFound(_)) => Ok(ct::OPC_RELATIONSHIPS.to_string()),
            Err(error) => Err(error),
        }
    }

    /// Return whether a ZIP member is an OPC relationship part rather than an
    /// ordinary package part. Relationship parts live in an `_rels` directory
    /// and have a `.rels` suffix.
    fn is_relationship_member(member_name: &str) -> bool {
        let Some((directory, filename)) = member_name.rsplit_once('/') else {
            return false;
        };
        let has_rels_extension = filename
            .rsplit_once('.')
            .is_some_and(|(_, extension)| extension.eq_ignore_ascii_case("rels"));
        has_rels_extension && (directory == "_rels" || directory.ends_with("/_rels"))
    }

    fn relationship_part_has_relationships(member_name: &str) -> bool {
        Self::is_relationship_member(member_name)
            && (member_name.starts_with("_rels/_rels/") || member_name.contains("/_rels/_rels/"))
    }

    /// Get an iterator over all serialized parts.
    pub fn iter_sparts(&self) -> impl Iterator<Item = &SerializedPart> {
        self.sparts.iter()
    }

    /// Get package-level relationships.
    #[must_use]
    pub fn pkg_srels(&self) -> &[SerializedRelationship] {
        &self.pkg_srels
    }

    /// Take ownership of package-level relationships (zero-copy move).
    pub fn take_pkg_srels(&mut self) -> SmallVec<[SerializedRelationship; 8]> {
        std::mem::take(&mut self.pkg_srels)
    }

    /// Take ownership of all serialized parts (zero-copy move).
    pub fn take_sparts(&mut self) -> Vec<SerializedPart> {
        std::mem::take(&mut self.sparts)
    }

    /// ZIP items that were present in the archive but are not OPC parts.
    ///
    /// Their contents are never decompressed; the entries exist so that
    /// tolerating archive junk does not hide it from the caller.
    #[must_use]
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Take ownership of the reported non-part members (zero-copy move).
    pub fn take_non_part_members(&mut self) -> Vec<NonPartMember> {
        std::mem::take(&mut self.non_part_members)
    }
}

#[derive(Default)]
struct RelationshipLedger {
    declared_xml_bytes: u64,
    retained_xml_bytes: u64,
    relationships: u64,
    xml_events: u64,
}

impl RelationshipLedger {
    fn preflight_xml_bytes(&mut self, limits: ReadLimits, bytes: u64) -> Result<()> {
        limits.check(
            ReadResource::RelationshipXmlBytes,
            bytes,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        self.declared_xml_bytes = checked_add(
            self.declared_xml_bytes,
            bytes,
            ReadResource::TotalRelationshipXmlBytes,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        Ok(())
    }

    fn retain_xml_bytes(&mut self, limits: ReadLimits, bytes: u64) -> Result<()> {
        limits.check(
            ReadResource::RelationshipXmlBytes,
            bytes,
            limits.max_relationship_xml_bytes() as u64,
        )?;
        self.retained_xml_bytes = checked_add(
            self.retained_xml_bytes,
            bytes,
            ReadResource::TotalRelationshipXmlBytes,
            limits.max_total_relationship_xml_bytes() as u64,
        )?;
        Ok(())
    }

    fn add_relationship(&mut self, limits: ReadLimits) -> Result<()> {
        self.relationships = checked_add(
            self.relationships,
            1,
            ReadResource::TotalRelationships,
            limits.max_total_relationships() as u64,
        )?;
        Ok(())
    }

    fn add_event(&mut self, limits: ReadLimits) -> Result<()> {
        self.xml_events = checked_add(
            self.xml_events,
            1,
            ReadResource::TotalRelationshipXmlEvents,
            limits.max_total_relationship_xml_events() as u64,
        )?;
        Ok(())
    }
}

#[allow(
    clippy::too_many_arguments,
    reason = "relationship element parsing threads the shared parse state; bundling it into a struct would not reduce complexity"
)]
fn inspect_relationship_element(
    relationships: &mut SmallVec<[SerializedRelationship; 8]>,
    ids: &mut HashSet<String>,
    resolved_namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
    base_uri: &str,
    source_uri: Option<&str>,
    depth: usize,
    root_seen: &mut bool,
    limits: ReadLimits,
    ledger: &mut RelationshipLedger,
) -> Result<()> {
    let correct_namespace = matches!(
        resolved_namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == namespace::OPC_RELATIONSHIPS.as_bytes()
    );
    if depth == 0 {
        if *root_seen {
            return Err(OpcError::InvalidRelationshipsManifest(
                "multiple root elements".to_string(),
            ));
        }
        *root_seen = true;
        if element.local_name().as_ref() != b"Relationships" || !correct_namespace {
            return Err(OpcError::InvalidRelationshipsManifest(
                "root must be Relationships in the OPC relationships namespace".to_string(),
            ));
        }
        return Ok(());
    }
    if depth != 1 || element.local_name().as_ref() != b"Relationship" || !correct_namespace {
        return Err(OpcError::InvalidRelationshipsManifest(
            "only direct Relationship children are permitted".to_string(),
        ));
    }
    let next_relationship_count = checked_increment(
        relationships.len(),
        limits.max_relationships_per_part(),
        ReadResource::RelationshipsPerPart,
    )?;

    let mut id_attribute = None;
    let mut type_attribute = None;
    let mut target_attribute = None;
    let mut target_mode = TargetMode::Internal;
    for attribute_result in element.attributes() {
        let attribute = attribute_result.map_err(|error| {
            OpcError::InvalidRelationshipsManifest(format!(
                "invalid Relationship attribute: {error}"
            ))
        })?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            attribute.value.as_ref().len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map(|value| value.to_string())
            .map_err(|error| {
                OpcError::InvalidRelationshipsManifest(format!(
                    "invalid Relationship attribute value: {error}"
                ))
            })?;
        limits.check(
            ReadResource::XmlAttributeBytes,
            value.len() as u64,
            limits.max_xml_attribute_bytes() as u64,
        )?;
        match attribute.key.as_ref() {
            b"Id" => id_attribute = Some(value),
            b"Type" => type_attribute = Some(value),
            b"Target" => target_attribute = Some(value),
            b"TargetMode" => target_mode = TargetMode::parse(&value)?,
            b"xmlns" => {},
            name if name.starts_with(b"xmlns:") => {},
            _ => {
                return Err(OpcError::InvalidRelationshipsManifest(
                    "unexpected Relationship attribute".to_string(),
                ));
            },
        }
    }
    let id = required_relationship_attribute(id_attribute, "Id")?;
    let relationship_type = required_relationship_attribute(type_attribute, "Type")?;
    let target = required_relationship_attribute(target_attribute, "Target")?;
    limits.check(
        ReadResource::RelationshipTargetBytes,
        target.len() as u64,
        limits.max_relationship_target_bytes() as u64,
    )?;
    if !is_xml_id(&id) {
        return Err(OpcError::InvalidRelationshipsManifest(format!(
            "relationship Id '{id}' is not an XML ID"
        )));
    }
    ids.try_reserve(1)
        .map_err(|source| allocation("OPC relationship IDs", source))?;
    if !ids.insert(id.clone()) {
        return Err(OpcError::DuplicateRelationshipId(id));
    }
    if relationship_type.chars().any(char::is_whitespace)
        || relationship_type.chars().any(char::is_control)
        || target.chars().any(char::is_control)
    {
        return Err(OpcError::InvalidRelationshipsManifest(
            "relationship Type or Target is not a valid URI reference".to_string(),
        ));
    }
    relationships
        .try_reserve(1)
        .map_err(|_| OpcError::CollectionAllocation {
            resource: "OPC relationships",
        })?;
    ledger.add_relationship(limits)?;
    debug_assert_eq!(next_relationship_count, relationships.len() + 1);
    relationships.push(SerializedRelationship {
        base_uri: base_uri.to_string(),
        source_uri: source_uri.map(str::to_string),
        r_id: id,
        reltype: relationship_type,
        target_ref: target,
        target_mode,
    });
    Ok(())
}

fn required_relationship_attribute(value: Option<String>, name: &str) -> Result<String> {
    value.filter(|text| !text.is_empty()).ok_or_else(|| {
        OpcError::InvalidRelationshipsManifest(format!(
            "Relationship requires a non-empty {name} attribute"
        ))
    })
}

fn is_xml_id(value: &str) -> bool {
    let mut characters = value.chars();
    characters.next().is_some_and(is_ncname_start) && characters.all(is_ncname_char)
}

fn is_ncname_start(character: char) -> bool {
    matches!(
        character,
        'A'..='Z'
            | '_'
            | 'a'..='z'
            | '\u{00c0}'..='\u{00d6}'
            | '\u{00d8}'..='\u{00f6}'
            | '\u{00f8}'..='\u{02ff}'
            | '\u{0370}'..='\u{037d}'
            | '\u{037f}'..='\u{1fff}'
            | '\u{200c}'..='\u{200d}'
            | '\u{2070}'..='\u{218f}'
            | '\u{2c00}'..='\u{2fef}'
            | '\u{3001}'..='\u{d7ff}'
            | '\u{f900}'..='\u{fdcf}'
            | '\u{fdf0}'..='\u{fffd}'
            | '\u{10000}'..='\u{effff}'
    )
}

fn is_ncname_char(character: char) -> bool {
    is_ncname_start(character)
        || matches!(
            character,
            '-' | '.' | '0'..='9' | '\u{00b7}' | '\u{0300}'..='\u{036f}' | '\u{203f}'..='\u{2040}'
        )
}

fn is_core_properties_relationship(relationship_type: &str) -> bool {
    matches!(
        relationship_type,
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
            | "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties"
    )
}

fn allocation(resource: &'static str, source: TryReserveError) -> OpcError {
    OpcError::Allocation { resource, source }
}

/// Whether a ZIP error reports a missing member rather than a damaged one.
fn is_member_missing(error: &soapberry_zip::Error) -> bool {
    matches!(error.kind(), soapberry_zip::ErrorKind::FileNotFound(_))
}

fn checked_increment(current: usize, maximum: usize, resource: ReadResource) -> Result<usize> {
    let actual = current.checked_add(1).ok_or(OpcError::ReadLimit {
        resource,
        actual: u64::MAX,
        maximum: maximum as u64,
    })?;
    if actual > maximum {
        return Err(OpcError::ReadLimit {
            resource,
            actual: actual as u64,
            maximum: maximum as u64,
        });
    }
    Ok(actual)
}

fn checked_add(current: u64, additional: u64, resource: ReadResource, maximum: u64) -> Result<u64> {
    let actual = current.checked_add(additional).ok_or(OpcError::ReadLimit {
        resource,
        actual: u64::MAX,
        maximum,
    })?;
    if actual > maximum {
        return Err(OpcError::ReadLimit {
            resource,
            actual,
            maximum,
        });
    }
    Ok(actual)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    fn package_bytes(root_relationships: &[u8], document: &[u8]) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        writer.write_stored("word/document.xml", document).unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn test_content_type_map() {
        let xml = br#"<?xml version="1.0"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                <Default Extension="xml" ContentType="application/xml"/>
                <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>"#;

        let ct_map = ContentTypeMap::from_xml(xml, ReadLimits::default()).unwrap();

        let uri = PackURI::new("/test.xml").unwrap();
        assert_eq!(ct_map.get(&uri).unwrap(), "application/xml");

        let document_uri = PackURI::new("/word/document.xml").unwrap();
        assert_eq!(
            ct_map.get(&document_uri).unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        );
    }

    #[test]
    fn missing_relationship_parts_are_optional_but_malformed_xml_is_not() {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer.write_stored("placeholder", b"x").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let archive = soapberry_zip::office::LazyArchiveReader::new(&bytes).unwrap();
        let package_uri = PackURI::new(PACKAGE_URI).unwrap();
        assert!(
            PackageReader::load_rels_lazy(
                &archive,
                &package_uri,
                ReadLimits::default(),
                &mut RelationshipLedger::default(),
            )
            .unwrap()
            .is_empty()
        );

        let malformed_bytes = package_bytes(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="document""#,
            b"document",
        );
        let malformed_archive =
            soapberry_zip::office::LazyArchiveReader::new(&malformed_bytes).unwrap();
        let error = PackageReader::load_rels_lazy(
            &malformed_archive,
            &package_uri,
            ReadLimits::default(),
            &mut RelationshipLedger::default(),
        )
        .unwrap_err();
        assert!(matches!(error, OpcError::QuickXmlError(_)));
    }

    #[test]
    fn corrupt_required_parts_report_zip_errors() {
        const DOCUMENT: &[u8] = b"unique required document payload";
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="word/document.xml"/></Relationships>"#;
        let mut bytes = package_bytes(relationships, DOCUMENT);
        let position = bytes
            .windows(DOCUMENT.len())
            .position(|window| window == DOCUMENT)
            .unwrap();
        bytes[position] ^= 0xff;

        let physical = PhysPkgReader::new(&bytes).unwrap();
        let Err(error) = PackageReader::from_phys_reader(&physical) else {
            panic!("corrupt required part unexpectedly loaded")
        };
        assert!(matches!(error, OpcError::ZipError(_)));
    }

    #[test]
    fn invalid_internal_relationship_targets_are_rejected() {
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="../escape.xml"/></Relationships>"#;
        let bytes = package_bytes(relationships, b"document");
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let Err(error) = PackageReader::from_phys_reader(&physical) else {
            panic!("invalid relationship target unexpectedly loaded")
        };
        assert!(matches!(error, OpcError::InvalidPackUri(_)));
    }

    fn package_with_physical_parts(first: &str, second: &str) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="gif" ContentType="image/gif"/></Types>"#,
            )
            .unwrap();
        writer.write_stored(first, b"first").unwrap();
        writer.write_stored(second, b"second").unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn rejects_equivalent_and_derived_physical_part_names() {
        let bytes = package_with_physical_parts("word/document.xml", "WORD/DOCUMENT.XML");
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let Err(error) = PackageReader::from_phys_reader(&physical) else {
            panic!("case-equivalent part names unexpectedly loaded")
        };
        assert!(matches!(error, OpcError::EquivalentPartNames { .. }));

        let derived_bytes =
            package_with_physical_parts("word/document.xml", "word/document.xml/image.gif");
        let derived_physical = PhysPkgReader::new(&derived_bytes).unwrap();
        let Err(derived_error) = PackageReader::from_phys_reader(&derived_physical) else {
            panic!("derived part names unexpectedly loaded")
        };
        assert!(matches!(derived_error, OpcError::DerivedPartNames { .. }));
    }

    #[test]
    fn rejects_apache_poi_derived_part_name_fixture() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j/OPCCompliance_DerivedPartNameFAIL.docx");
        let bytes = std::fs::read(path).unwrap();
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let Err(error) = PackageReader::from_phys_reader(&physical) else {
            panic!("Apache POI derived-name failure fixture unexpectedly loaded")
        };
        assert!(matches!(error, OpcError::DerivedPartNames { .. }));
    }

    fn relationships_xml(children: &str) -> String {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{children}</Relationships>"#
        )
    }

    #[test]
    fn relationships_manifest_requires_schema_root_and_direct_children() {
        for xml in [
            "<Relationships/>",
            r#"<Relationships xmlns="urn:wrong"/>"#,
            r#"<Wrong xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            &relationships_xml(
                r#"<Wrapper><Relationship Id="rId1" Type="urn:test" Target="part.xml"/></Wrapper>"#,
            ),
        ] {
            assert!(matches!(
                PackageReader::parse_rels_xml(xml.as_bytes(), "/word"),
                Err(OpcError::InvalidRelationshipsManifest(_))
            ));
        }
    }

    #[test]
    fn relationships_require_attributes_ids_and_exact_target_modes() {
        for child in [
            r#"<Relationship Type="urn:test" Target="part.xml"/>"#,
            r#"<Relationship Id="rId1" Target="part.xml"/>"#,
            r#"<Relationship Id="rId1" Type="urn:test"/>"#,
            r#"<Relationship Id="1bad" Type="urn:test" Target="part.xml"/>"#,
            r#"<Relationship Id="r:id" Type="urn:test" Target="part.xml"/>"#,
            r#"<Relationship Id="rId1" Type="urn:test" Target="part.xml" TargetMode="external"/>"#,
        ] {
            let xml = relationships_xml(child);
            assert!(PackageReader::parse_rels_xml(xml.as_bytes(), "/word").is_err());
        }

        let duplicate = relationships_xml(
            r#"<Relationship Id="rId1" Type="urn:a" Target="a.xml"/><Relationship Id="rId1" Type="urn:b" Target="b.xml"/>"#,
        );
        assert!(matches!(
            PackageReader::parse_rels_xml(duplicate.as_bytes(), "/word"),
            Err(OpcError::DuplicateRelationshipId(id)) if id == "rId1"
        ));
    }

    fn read_with_limits(bytes: &[u8], limits: ReadLimits) -> Result<PackageReader> {
        let physical = PhysPkgReader::new_with_limits(bytes, limits)?;
        PackageReader::from_phys_reader(&physical)
    }

    fn relationship(id: &str, relationship_type: &str, target: &str) -> String {
        format!(r#"<Relationship Id="{id}" Type="{relationship_type}" Target="{target}"/>"#)
    }

    fn package_with_extra_parts(
        root_relationships: &[u8],
        parts: &[(&str, &[u8])],
        document_relationships: Option<&[u8]>,
    ) -> Vec<u8> {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer
            .write_stored("_rels/.rels", root_relationships)
            .unwrap();
        for (name, data) in parts {
            writer.write_stored(name, data).unwrap();
        }
        if let Some(rels) = document_relationships {
            writer
                .write_stored("word/_rels/document.xml.rels", rels)
                .unwrap();
        }
        writer.finish_to_bytes().unwrap()
    }

    fn assert_limit(bytes: &[u8], limits: ReadLimits, resource: ReadResource) {
        match read_with_limits(bytes, limits) {
            Err(OpcError::ReadLimit {
                resource: actual, ..
            }) if actual == resource => {},
            Err(error) => panic!("expected {resource} limit, got {error:?}"),
            Ok(_) => panic!("expected {resource} limit, package was accepted"),
        }
    }

    #[test]
    fn read_limits_bound_physical_and_retained_parts_before_parallel_loading() {
        let root = relationships_xml(&relationship("rId1", "urn:test", "word/document.xml"));
        let exact =
            package_with_extra_parts(root.as_bytes(), &[("word/document.xml", b"ok")], None);
        assert!(
            read_with_limits(
                &exact,
                ReadLimits::builder()
                    .max_archive_members(3)
                    .unwrap()
                    .max_relationship_parts(1)
                    .unwrap()
                    .max_parts(1)
                    .unwrap()
                    .max_part_bytes(2)
                    .unwrap()
                    .max_total_part_bytes(2)
                    .unwrap()
                    .build()
                    .unwrap(),
            )
            .is_ok()
        );

        assert_limit(
            &exact,
            ReadLimits::builder()
                .max_archive_members(2)
                .unwrap()
                .max_relationship_parts(1)
                .unwrap()
                .max_parts(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::ArchiveMembers,
        );
        assert_limit(
            &package_with_extra_parts(
                root.as_bytes(),
                &[("word/document.xml", b"ok"), ("word/extra.xml", b"ok")],
                None,
            ),
            ReadLimits::builder().max_parts(1).unwrap().build().unwrap(),
            ReadResource::Parts,
        );
        assert_limit(
            &exact,
            ReadLimits::builder()
                .max_part_bytes(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::PartBytes,
        );
        assert_limit(
            &package_with_extra_parts(
                root.as_bytes(),
                &[("word/document.xml", b"ok"), ("word/extra.xml", b"ok")],
                None,
            ),
            ReadLimits::builder()
                .max_total_part_bytes(3)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::TotalPartBytes,
        );
    }

    #[test]
    fn read_limits_bound_relationship_resources_at_exact_boundaries() {
        const RELATIONSHIPS_CONTENT_TYPE: &str =
            "application/vnd.openxmlformats-package.relationships+xml";
        let internal = relationship("rId1", "urn:x", "word/document.xml");
        let root = relationships_xml(&internal);
        let exact = package_with_extra_parts(root.as_bytes(), &[("word/document.xml", b"x")], None);
        let exact_limits = ReadLimits::builder()
            .max_relationship_parts(1)
            .unwrap()
            .max_relationship_xml_bytes(root.len())
            .unwrap()
            .max_total_relationship_xml_bytes(root.len())
            .unwrap()
            .max_relationships_per_part(1)
            .unwrap()
            .max_total_relationships(1)
            .unwrap()
            .max_relationship_graph_nodes(1)
            .unwrap()
            .max_xml_events(5)
            .unwrap()
            .max_total_relationship_xml_events(4)
            .unwrap()
            .max_xml_depth(1)
            .unwrap()
            .max_xml_attribute_bytes(RELATIONSHIPS_CONTENT_TYPE.len())
            .unwrap()
            .max_relationship_target_bytes("/word/document.xml".len())
            .unwrap()
            .build()
            .unwrap();
        assert!(read_with_limits(&exact, exact_limits).is_ok());

        let child = relationships_xml(&relationship("rId2", "urn:x", "https://x"));
        let two_relationship_parts = package_with_extra_parts(
            root.as_bytes(),
            &[("word/document.xml", b"x")],
            Some(child.as_bytes()),
        );
        assert_limit(
            &two_relationship_parts,
            ReadLimits::builder()
                .max_relationship_parts(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipParts,
        );
        assert_limit(
            &exact,
            ReadLimits::builder()
                .max_relationship_xml_bytes(root.len() - 1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipXmlBytes,
        );
        assert_limit(
            &two_relationship_parts,
            ReadLimits::builder()
                .max_total_relationship_xml_bytes(root.len() + child.len() - 1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::TotalRelationshipXmlBytes,
        );

        let two_relationships = relationships_xml(&format!(
            "{}{}",
            internal,
            relationship("rId2", "urn:x", "other.xml"),
        ));
        let two_edges = package_with_extra_parts(
            two_relationships.as_bytes(),
            &[("word/document.xml", b"x")],
            None,
        );
        assert_limit(
            &two_edges,
            ReadLimits::builder()
                .max_relationships_per_part(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipsPerPart,
        );
        assert_limit(
            &two_edges,
            ReadLimits::builder()
                .max_relationships_per_part(2)
                .unwrap()
                .max_total_relationships(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::TotalRelationships,
        );
        assert_limit(
            &two_edges,
            ReadLimits::builder()
                .max_relationship_graph_nodes(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipGraphNodes,
        );

        let event_bomb = relationships_xml(&format!("{internal}<!--a--><!--b-->"));
        assert_limit(
            &package_with_extra_parts(event_bomb.as_bytes(), &[("word/document.xml", b"x")], None),
            ReadLimits::builder()
                .max_xml_events(5)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::XmlEvents,
        );
        assert_limit(
            &two_relationship_parts,
            ReadLimits::builder()
                .max_total_relationship_xml_events(7)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::TotalRelationshipXmlEvents,
        );

        let nested = relationships_xml("<Wrapper></Wrapper>");
        assert_limit(
            &package_with_extra_parts(nested.as_bytes(), &[("word/document.xml", b"x")], None),
            ReadLimits::builder()
                .max_xml_depth(1)
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::XmlDepth,
        );
        let encoded_type = format!("urn:{}", "&#x78;".repeat(10));
        let long_attribute = relationships_xml(&relationship("rId1", &encoded_type, "a"));
        assert_limit(
            &package_with_extra_parts(
                long_attribute.as_bytes(),
                &[("word/document.xml", b"x")],
                None,
            ),
            ReadLimits::builder()
                .max_xml_attribute_bytes(RELATIONSHIPS_CONTENT_TYPE.len())
                .unwrap()
                .max_relationship_target_bytes(RELATIONSHIPS_CONTENT_TYPE.len())
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::XmlAttributeBytes,
        );
        let external = relationships_xml(
            r#"<Relationship Id="rId1" Type="urn:x" Target="https://x" TargetMode="External"/>"#,
        );
        assert_limit(
            &package_with_extra_parts(external.as_bytes(), &[("word/document.xml", b"x")], None),
            ReadLimits::builder()
                .max_relationship_target_bytes("https://".len())
                .unwrap()
                .build()
                .unwrap(),
            ReadResource::RelationshipTargetBytes,
        );
    }

    #[test]
    fn relationship_targets_preserve_fragments_and_typed_modes() {
        let xml = relationships_xml(
            r#"<Relationship Id="rId1" Type="urn:internal" Target="media/image.png?size=2#preview"/><Relationship Id="external" Type="urn:external" Target="mailto:dev@example.test?subject=Hi#body" TargetMode="External"/>"#,
        );
        let relationships = PackageReader::parse_rels_xml(xml.as_bytes(), "/word").unwrap();
        assert_eq!(relationships[0].target_mode, TargetMode::Internal);
        assert_eq!(
            relationships[0].target_partname().unwrap().as_str(),
            "/word/media/image.png"
        );
        assert_eq!(
            relationships[0].target_ref,
            "media/image.png?size=2#preview"
        );
        assert_eq!(relationships[1].target_mode, TargetMode::External);
        assert_eq!(
            relationships[1].target_ref,
            "mailto:dev@example.test?subject=Hi#body"
        );
    }

    #[test]
    fn package_rejects_multiple_core_properties_relationships() {
        let xml = relationships_xml(
            r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties" Target="docProps/core2.xml"/>"#,
        );
        assert!(matches!(
            PackageReader::parse_rels_xml(xml.as_bytes(), "/"),
            Err(OpcError::MultipleCorePropertiesRelationships)
        ));
    }

    #[test]
    fn rejects_poi_relationships_entity_fixture() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j/PackageRelsHasEntities.ooxml");
        let bytes = std::fs::read(path).unwrap();
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let Err(error) = PackageReader::from_phys_reader(&physical) else {
            panic!("entity-bearing relationships fixture unexpectedly loaded")
        };
        assert!(matches!(error, OpcError::InvalidRelationshipsManifest(_)));
    }

    #[test]
    fn preserves_poi_special_relationship_targets_across_round_trip() {
        fn targets(package: &crate::OpcPackage) -> Vec<String> {
            let mut values: Vec<String> = package
                .rels()
                .iter()
                .chain(package.iter_parts().flat_map(|part| part.rels().iter()))
                .map(|relationship| relationship.target_ref().to_string())
                .collect();
            values.sort();
            values
        }

        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j/50154.xlsx");
        let bytes = std::fs::read(path).unwrap();
        let package = crate::OpcPackage::from_bytes(&bytes).unwrap();
        let original_targets = targets(&package);
        assert!(original_targets.iter().any(|target| target.contains('#')));
        assert!(
            original_targets
                .iter()
                .any(|target| target.contains("Another Sheet"))
        );

        let serialized = crate::PackageWriter::to_bytes(&package).unwrap();
        let reloaded = crate::OpcPackage::from_bytes(&serialized).unwrap();
        assert_eq!(targets(&reloaded), original_targets);
    }
}
