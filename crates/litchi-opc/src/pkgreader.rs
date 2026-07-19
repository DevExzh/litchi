//! Low-level, read-only API to a serialized Open Packaging Convention (OPC) package.
//!
//! This module provides the PackageReader for parsing OPC packages, including
//! content type mapping, relationship resolution, and part loading. It uses
//! efficient algorithms for parsing and minimal memory allocation.

use crate::constants::namespace;
use crate::content_type::ContentTypeMap;
use crate::error::{OpcError, Result};
use crate::packuri::{PACKAGE_URI, PackURI, PartNameConflict};
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
        let content_types_path = crate::packuri::CONTENT_TYPES_URI.trim_start_matches('/');
        let content_types_xml = archive
            .read(content_types_path)
            .map_err(|_| OpcError::PartNotFound("[Content_Types].xml".to_string()))?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml)?;

        // Phase 2: Get package-level relationships (on-demand decompression)
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels = Self::load_rels_lazy(archive, &package_uri)?;

        // Phase 3: Load every physical part. Relationship traversal alone is
        // insufficient because OPC permits parts with no incoming relationship.
        let sparts = Self::load_parts_lazy(archive, &pkg_srels, &content_types)?;

        Ok(Self { pkg_srels, sparts })
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
            Err(error) if matches!(error.kind(), soapberry_zip::ErrorKind::FileNotFound(_)) => {
                Ok(SmallVec::new())
            },
            Err(error) => Err(error.into()),
        }
    }

    /// Load all physical parts using parallel decompression.
    ///
    /// This is a two-phase approach for maximum performance:
    /// 1. Traverse relationships to validate targets and discover their source parts.
    /// 2. Enumerate ZIP members to include parts with no incoming relationship.
    /// 3. Decompress all ordinary part contents in parallel.
    fn load_parts_lazy(
        archive: &soapberry_zip::office::LazyArchiveReader<'_>,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
    ) -> Result<Vec<SerializedPart>> {
        use std::collections::HashSet;

        // Phase 1: Discover relationship-reachable parts. Relationship types
        // belong to edges, not parts, so they are intentionally not stored here.
        let mut discovered: Vec<(PackURI, SmallVec<[SerializedRelationship; 8]>)> =
            Vec::with_capacity(32);
        let mut visited = HashSet::with_capacity(32);
        let mut work_queue: Vec<PackURI> = Vec::with_capacity(pkg_srels.len());

        // Initialize work queue with package-level relationships
        for srel in pkg_srels {
            if srel.is_external() {
                continue;
            }
            let partname = srel.target_partname()?;
            let partname_str = partname.to_string();
            if visited.insert(partname_str) {
                work_queue.push(partname);
            }
        }

        // Traverse relationship graph (only decompresses small .rels files)
        while let Some(partname) = work_queue.pop() {
            // Load relationships for this part
            let part_srels = Self::load_rels_lazy(archive, &partname)?;

            // Add child parts to work queue
            for child_srel in &part_srels {
                if child_srel.is_external() {
                    continue;
                }
                let child_partname = child_srel.target_partname()?;
                let child_partname_str = child_partname.to_string();
                if visited.insert(child_partname_str) {
                    work_queue.push(child_partname);
                }
            }

            discovered.push((partname, part_srels));
        }

        // Phase 2: Validate the physical part-name collection before reading any
        // potentially large payload. Relationship parts participate in OPC name
        // conformance even though they are loaded separately.
        let mut physical_names: Vec<PackURI> = Vec::new();
        for member_name in archive.file_names() {
            if member_name.is_empty()
                || member_name.ends_with('/')
                || member_name == crate::packuri::CONTENT_TYPES_URI.trim_start_matches('/')
            {
                continue;
            }

            if Self::relationship_part_has_relationships(member_name) {
                return Err(OpcError::RelationshipPartCannotBeSource(
                    member_name.to_string(),
                ));
            }

            let absolute_name = format!("/{member_name}");
            let partname = PackURI::new(&absolute_name).map_err(OpcError::InvalidPackUri)?;
            for existing in &physical_names {
                if let Some(conflict) = existing.conflict_with(&partname) {
                    return Err(part_name_conflict_error(existing, &partname, conflict));
                }
            }
            for (reachable, _) in &discovered {
                if reachable != &partname
                    && let Some(conflict) = reachable.conflict_with(&partname)
                {
                    return Err(part_name_conflict_error(reachable, &partname, conflict));
                }
            }
            physical_names.push(partname);
        }

        // Enumerate ordinary physical members so unreferenced custom data and
        // extension parts are retained during open/save round trips.
        for partname in physical_names {
            if Self::is_relationship_member(partname.membername()) {
                content_types.get(&partname)?;
                continue;
            }
            if visited.insert(partname.to_string()) {
                let part_srels = Self::load_rels_lazy(archive, &partname)?;
                discovered.push((partname, part_srels));
            }
        }

        // Resolve content types before decompressing potentially large parts.
        let mut typed_parts = Vec::with_capacity(discovered.len());
        for (partname, part_srels) in discovered {
            let content_type = content_types.get(&partname)?;
            typed_parts.push((partname, content_type, part_srels));
        }

        // Phase 3: Parallel decompression of all discovered parts
        // Collect member names for parallel batch read
        let member_names: Vec<&str> = typed_parts
            .iter()
            .map(|(partname, _, _)| partname.membername())
            .collect();

        // Decompress all part contents in parallel
        let mut decompressed = HashMap::with_capacity(member_names.len());
        for (member_name, result) in archive.read_many_parallel_results(&member_names) {
            decompressed.insert(member_name.to_string(), result?);
        }

        // Phase 4: Build SerializedPart structures (take ownership, no cloning)
        let mut sparts = Vec::with_capacity(typed_parts.len());
        for (partname, content_type, part_srels) in typed_parts {
            let membername = partname.membername();
            // Remove from map to take ownership instead of cloning
            let blob = decompressed
                .remove(membername)
                .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))?;
            sparts.push(SerializedPart {
                partname,
                content_type,
                blob,
                srels: part_srels,
            });
        }

        Ok(sparts)
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
            .join("../../3rdparty/poi/test-data/openxml4j/OPCCompliance_DerivedPartNameFAIL.docx");
        let bytes = std::fs::read(path).unwrap();
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("Apache POI derived-name failure fixture unexpectedly loaded"),
            Err(error) => error,
        };
        assert!(matches!(error, OpcError::DerivedPartNames { .. }));
    }

    #[test]
    fn rejects_unreferenced_part_without_content_type_mapping() {
        let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
        writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
        writer.write_stored("custom/orphan.bin", b"orphan").unwrap();
        let bytes = writer.finish_to_bytes().unwrap();
        let physical = PhysPkgReader::new(&bytes).unwrap();
        let error = match PackageReader::from_phys_reader(&physical) {
            Ok(_) => panic!("unmapped orphan part unexpectedly loaded"),
            Err(error) => error,
        };
        assert!(matches!(error, OpcError::ContentTypeNotFound(_)));
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
            .join("../../3rdparty/poi/test-data/openxml4j/PackageRelsHasEntities.ooxml");
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
            .join("../../3rdparty/poi/test-data/openxml4j/50154.xlsx");
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
