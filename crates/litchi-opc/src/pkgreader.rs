//! Low-level, read-only API to a serialized Open Packaging Convention (OPC) package.
//!
//! This module provides the PackageReader for parsing OPC packages, including
//! content type mapping, relationship resolution, and part loading. It uses
//! efficient algorithms for parsing and minimal memory allocation.

use crate::constants::{content_type as ct, namespace};
use crate::content_type::ContentTypeMap;
use crate::error::{OpcError, Result};
use crate::members::{NonPartMember, NonPartReason, PartNameIndex, part_name_for_member};
use crate::packuri::{PACKAGE_URI, PackURI};
use crate::phys_pkg::PhysPkgReader;
use crate::rel::{TargetMode, relationship_target_components};
use quick_xml::XmlVersion;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::collections::HashMap;

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
    /// Uses SmallVec for efficient storage of typically small relationship collections
    pub srels: SmallVec<[SerializedRelationship; 8]>,
}

#[cfg(test)]
mod physical_part_tests {
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
        let orphan = reloaded
            .iter_parts()
            .find(|part| part.partname().as_str() == ORPHAN_PART_NAME)
            .expect("unreferenced physical part must survive save and reopen");
        assert_eq!(orphan.blob(), ORPHAN_CONTENT);
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
    pub fn is_external(&self) -> bool {
        self.target_mode == TargetMode::External
    }

    /// Get the target partname for internal relationships.
    ///
    /// Resolves the relative target reference against the base URI
    /// to produce an absolute PackURI.
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
    /// Uses SmallVec for efficient storage of typically small relationship collections
    pkg_srels: SmallVec<[SerializedRelationship; 8]>,

    /// All serialized parts in the package
    sparts: Vec<SerializedPart>,

    /// ZIP items that were present but are not OPC parts
    non_part_members: Vec<NonPartMember>,
}

/// Reserved ZIP item name of the content types stream (ECMA-376 Part 2 §10.1.2.2).
const CONTENT_TYPES_MEMBER: &str = "[Content_Types].xml";

/// Upper bound on distinct part names visited while walking the relationship graph.
///
/// Relationship targets need not resolve to an existing ZIP item, so a hostile
/// package could otherwise name an unbounded number of phantom parts and make
/// the traversal run for as long as its relationship manifests allow.
const MAX_RELATIONSHIP_GRAPH_NODES: usize = 100_000;

/// Whether a ZIP error reports a missing member rather than a damaged one.
fn is_member_missing(error: &soapberry_zip::Error) -> bool {
    matches!(error.kind(), soapberry_zip::ErrorKind::FileNotFound(_))
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
    /// A new PackageReader with all parts and relationships loaded
    pub fn from_phys_reader(phys_reader: &PhysPkgReader<'_>) -> Result<Self> {
        let archive = phys_reader.archive();

        // Phase 1: Decompress and parse content types (on-demand)
        let content_types_member = Self::locate_content_types_member(archive)?;
        let content_types_xml = archive.read(content_types_member)?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml)?;

        // Phase 2: Get package-level relationships (on-demand decompression)
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels = Self::load_rels_lazy(archive, &package_uri)?;

        // Phase 3: Load every physical part. Relationship traversal alone is
        // insufficient because OPC permits parts with no incoming relationship.
        let mut non_part_members = Vec::new();
        let sparts = Self::load_parts_lazy(
            archive,
            content_types_member,
            &pkg_srels,
            &content_types,
            &mut non_part_members,
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
    fn locate_content_types_member<'archive>(
        archive: &'archive soapberry_zip::office::LazyArchiveReader<'_>,
    ) -> Result<&'archive str> {
        if archive.contains(CONTENT_TYPES_MEMBER) {
            return Ok(CONTENT_TYPES_MEMBER);
        }
        archive
            .file_names()
            .find(|name| name.eq_ignore_ascii_case(CONTENT_TYPES_MEMBER))
            .ok_or_else(|| OpcError::PartNotFound(CONTENT_TYPES_MEMBER.to_string()))
    }

    /// Parse relationships XML into SerializedRelationship structs.
    #[cfg(test)]
    fn parse_rels_xml(
        rels_xml: &[u8],
        base_uri: &str,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        Self::parse_rels_xml_with_source(rels_xml, base_uri, None)
    }

    fn parse_rels_xml_with_source(
        rels_xml: &[u8],
        base_uri: &str,
        source_uri: Option<&str>,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let mut srels = SmallVec::new();
        let mut reader = NsReader::from_reader(rels_xml);
        reader.config_mut().trim_text(true);
        let mut depth = 0usize;
        let mut root_seen = false;
        let mut ids = std::collections::HashSet::new();

        loop {
            let decoder = reader.decoder();
            let (resolved_namespace, event) = reader.read_resolved_event().map_err(|error| {
                OpcError::InvalidRelationshipsManifest(format!("XML parse error: {error}"))
            })?;
            match event {
                Event::Start(element) => {
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
                    )?;
                    depth = depth.checked_add(1).ok_or_else(|| {
                        OpcError::InvalidRelationshipsManifest(
                            "XML nesting depth overflow".to_string(),
                        )
                    })?;
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
                _ => {},
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
    fn load_rels_lazy(
        archive: &soapberry_zip::office::LazyArchiveReader<'_>,
        source_uri: &PackURI,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let rels_uri = source_uri.rels_uri().map_err(OpcError::InvalidPackUri)?;
        let rels_path = rels_uri.membername();

        match archive.read(rels_path) {
            Ok(rels_xml) => Self::parse_rels_xml_with_source(
                &rels_xml,
                source_uri.base_uri(),
                Some(source_uri.as_str()),
            ),
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
    fn load_parts_lazy(
        archive: &soapberry_zip::office::LazyArchiveReader<'_>,
        content_types_member: &str,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
        non_part_members: &mut Vec<NonPartMember>,
    ) -> Result<Vec<SerializedPart>> {
        // Phase 1: relationship reachability. Relationship types belong to
        // edges, not parts, so they are intentionally not recorded here.
        let mut relationships = Self::walk_relationship_graph(archive, pkg_srels)?;

        // Phase 2 and 3: classify members, then admit the survivors as parts.
        let mut index = PartNameIndex::with_capacity(archive.len());
        let mut typed_parts: Vec<(PackURI, String)> = Vec::with_capacity(archive.len());
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

            let Some(partname) = part_name_for_member(member_name) else {
                non_part_members.push(NonPartMember::new(
                    member_name,
                    NonPartReason::UnmappablePartName,
                ));
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
                        ));
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
                typed_parts.push((partname, content_type));
            }
        }

        // Phase 4: parallel decompression of every admitted part.
        let member_names: Vec<&str> = typed_parts
            .iter()
            .map(|(partname, _)| partname.membername())
            .collect();
        let mut decompressed = HashMap::with_capacity(member_names.len());
        for (member_name, result) in archive.read_many_parallel_results(&member_names) {
            decompressed.insert(member_name.to_string(), result?);
        }

        // Phase 5: build SerializedPart structures (take ownership, no cloning)
        let mut sparts = Vec::with_capacity(typed_parts.len());
        for (partname, content_type) in typed_parts {
            let srels = match relationships.remove(partname.as_str()) {
                Some(srels) => srels,
                None => Self::load_rels_lazy(archive, &partname)?,
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

    /// Walk the relationship graph, returning each visited part name with its
    /// own relationships.
    ///
    /// Only small `.rels` members are decompressed. A target that has no ZIP
    /// member is still visited: OPC defines no rule requiring a relationship
    /// target to resolve, and the dangling relationship stays visible on its
    /// source part.
    fn walk_relationship_graph(
        archive: &soapberry_zip::office::LazyArchiveReader<'_>,
        pkg_srels: &[SerializedRelationship],
    ) -> Result<HashMap<String, SmallVec<[SerializedRelationship; 8]>>> {
        let mut visited: HashMap<String, SmallVec<[SerializedRelationship; 8]>> = HashMap::new();
        let mut work_queue: Vec<PackURI> = Vec::with_capacity(pkg_srels.len());
        for srel in pkg_srels {
            Self::enqueue_target(srel, &mut visited, &mut work_queue)?;
        }

        while let Some(partname) = work_queue.pop() {
            let part_srels = Self::load_rels_lazy(archive, &partname)?;
            for child_srel in &part_srels {
                Self::enqueue_target(child_srel, &mut visited, &mut work_queue)?;
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
    ) -> Result<()> {
        if srel.is_external() {
            return Ok(());
        }
        let partname = srel.target_partname()?;
        if visited.contains_key(partname.as_str()) {
            return Ok(());
        }
        if visited.len() >= MAX_RELATIONSHIP_GRAPH_NODES {
            return Err(OpcError::InvalidRelationshipsManifest(format!(
                "package refers to more than {MAX_RELATIONSHIP_GRAPH_NODES} distinct part names"
            )));
        }
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
        filename.ends_with(".rels") && (directory == "_rels" || directory.ends_with("/_rels"))
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
    pub fn non_part_members(&self) -> &[NonPartMember] {
        &self.non_part_members
    }

    /// Take ownership of the reported non-part members (zero-copy move).
    pub fn take_non_part_members(&mut self) -> Vec<NonPartMember> {
        std::mem::take(&mut self.non_part_members)
    }
}

const MAX_RELATIONSHIPS_PER_PART: usize = 65_536;

fn inspect_relationship_element(
    relationships: &mut SmallVec<[SerializedRelationship; 8]>,
    ids: &mut std::collections::HashSet<String>,
    resolved_namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
    base_uri: &str,
    source_uri: Option<&str>,
    depth: usize,
    root_seen: &mut bool,
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
    if relationships.len() >= MAX_RELATIONSHIPS_PER_PART {
        return Err(OpcError::InvalidRelationshipsManifest(format!(
            "relationships part exceeds {MAX_RELATIONSHIPS_PER_PART} entries"
        )));
    }

    let mut id = None;
    let mut relationship_type = None;
    let mut target = None;
    let mut target_mode = TargetMode::Internal;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            OpcError::InvalidRelationshipsManifest(format!(
                "invalid Relationship attribute: {error}"
            ))
        })?;
        let value = || {
            attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map(|value| value.to_string())
        };
        match attribute.key.as_ref() {
            b"Id" => id = Some(value()?),
            b"Type" => relationship_type = Some(value()?),
            b"Target" => target = Some(value()?),
            b"TargetMode" => target_mode = TargetMode::parse(&value()?)?,
            b"xmlns" => {},
            name if name.starts_with(b"xmlns:") => {},
            _ => {
                return Err(OpcError::InvalidRelationshipsManifest(
                    "unexpected Relationship attribute".to_string(),
                ));
            },
        }
    }
    let id = required_relationship_attribute(id, "Id")?;
    let relationship_type = required_relationship_attribute(relationship_type, "Type")?;
    let target = required_relationship_attribute(target, "Target")?;
    if !is_xml_id(&id) {
        return Err(OpcError::InvalidRelationshipsManifest(format!(
            "relationship Id '{id}' is not an XML ID"
        )));
    }
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
    value.filter(|value| !value.is_empty()).ok_or_else(|| {
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

#[cfg(test)]
mod tests {
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

        let ct_map = ContentTypeMap::from_xml(xml).unwrap();

        let uri = PackURI::new("/test.xml").unwrap();
        assert_eq!(ct_map.get(&uri).unwrap(), "application/xml");

        let uri = PackURI::new("/word/document.xml").unwrap();
        assert_eq!(
            ct_map.get(&uri).unwrap(),
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
            PackageReader::load_rels_lazy(&archive, &package_uri)
                .unwrap()
                .is_empty()
        );

        let bytes = package_bytes(b"<Relationships><Relationship", b"document");
        let archive = soapberry_zip::office::LazyArchiveReader::new(&bytes).unwrap();
        let error = PackageReader::load_rels_lazy(&archive, &package_uri).unwrap_err();
        assert!(matches!(error, OpcError::InvalidRelationshipsManifest(_)));
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
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("corrupt required part unexpectedly loaded"),
            Err(error) => error,
        };
        assert!(matches!(error, OpcError::ZipError(_)));
    }

    #[test]
    fn invalid_internal_relationship_targets_are_rejected() {
        let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="officeDocument" Target="../escape.xml"/></Relationships>"#;
        let bytes = package_bytes(relationships, b"document");
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("invalid relationship target unexpectedly loaded"),
            Err(error) => error,
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
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("case-equivalent part names unexpectedly loaded"),
            Err(error) => error,
        };
        assert!(matches!(error, OpcError::EquivalentPartNames { .. }));

        let bytes = package_with_physical_parts("word/document.xml", "word/document.xml/image.gif");
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("derived part names unexpectedly loaded"),
            Err(error) => error,
        };
        assert!(matches!(error, OpcError::DerivedPartNames { .. }));
    }

    #[test]
    fn rejects_apache_poi_derived_part_name_fixture() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/openxml4j/OPCCompliance_DerivedPartNameFAIL.docx");
        let bytes = std::fs::read(path).unwrap();
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("Apache POI derived-name failure fixture unexpectedly loaded"),
            Err(error) => error,
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
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("entity-bearing relationships fixture unexpectedly loaded"),
            Err(error) => error,
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
