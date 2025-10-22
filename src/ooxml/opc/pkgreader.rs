use crate::ooxml::opc::constants::target_mode;
use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::{PACKAGE_URI, PackURI};
use crate::ooxml::opc::phys_pkg::PhysPkgReader;
use quick_xml::Reader;
use quick_xml::events::Event;
/// Low-level, read-only API to a serialized Open Packaging Convention (OPC) package.
///
/// This module provides the PackageReader for parsing OPC packages, including
/// content type mapping, relationship resolution, and part loading. It uses
/// efficient algorithms for parsing and minimal memory allocation.
use smallvec::SmallVec;
use std::collections::HashMap;
use std::io::{Read, Seek};

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

    /// The relationship type that refers to this part
    pub reltype: String,

    /// The binary content of this part
    pub blob: Vec<u8>,

    /// Serialized relationships from this part
    /// Uses SmallVec for efficient storage of typically small relationship collections
    pub srels: SmallVec<[SerializedRelationship; 8]>,
}

/// Serialized relationship as read from a .rels file.
///
/// Contains all relationship information in string form, before
/// being converted into Relationship objects with resolved part references.
#[derive(Debug, Clone)]
pub struct SerializedRelationship {
    /// Base URI for resolving relative references
    pub base_uri: String,

    /// Relationship ID (e.g., "rId1")
    pub r_id: String,

    /// Relationship type URI
    pub reltype: String,

    /// Target reference (relative URI or external URL)
    pub target_ref: String,

    /// Target mode (Internal or External)
    pub target_mode: String,
}

impl SerializedRelationship {
    /// Check if this is an external relationship.
    #[inline]
    pub fn is_external(&self) -> bool {
        self.target_mode == target_mode::EXTERNAL
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
        PackURI::from_rel_ref(&self.base_uri, &self.target_ref).map_err(OpcError::InvalidPackUri)
    }
}

/// Content type map for looking up content types by part name or extension.
///
/// Implements the OPC content type discovery algorithm using Default and Override elements
/// from [Content_Types].xml. Uses efficient hash maps for O(1) lookup.
struct ContentTypeMap {
    /// Maps file extensions to default content types
    defaults: HashMap<String, String>,

    /// Maps specific partnames to override content types
    overrides: HashMap<String, String>,
}

impl ContentTypeMap {
    /// Create a new empty content type map.
    fn new() -> Self {
        Self {
            defaults: HashMap::new(),
            overrides: HashMap::new(),
        }
    }

    /// Parse content types from [Content_Types].xml.
    ///
    /// Uses quick-xml for efficient streaming XML parsing with minimal allocation.
    fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut map = Self::new();
        let mut reader = Reader::from_reader(xml);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) => {
                    match e.local_name().as_ref() {
                        b"Default" => {
                            // Parse Default element: <Default Extension="xml" ContentType="application/xml"/>
                            let mut extension = None;
                            let mut content_type = None;

                            for attr in e.attributes() {
                                let attr = attr?;
                                match attr.key.as_ref() {
                                    b"Extension" => {
                                        extension = Some(attr.unescape_value()?.to_string());
                                    },
                                    b"ContentType" => {
                                        content_type = Some(attr.unescape_value()?.to_string());
                                    },
                                    _ => {},
                                }
                            }

                            if let (Some(ext), Some(ct)) = (extension, content_type) {
                                map.add_default(ext, ct);
                            }
                        },
                        b"Override" => {
                            // Parse Override element: <Override PartName="/word/document.xml" ContentType="..."/>
                            let mut partname = None;
                            let mut content_type = None;

                            for attr in e.attributes() {
                                let attr = attr?;
                                match attr.key.as_ref() {
                                    b"PartName" => {
                                        partname = Some(attr.unescape_value()?.to_string());
                                    },
                                    b"ContentType" => {
                                        content_type = Some(attr.unescape_value()?.to_string());
                                    },
                                    _ => {},
                                }
                            }

                            if let (Some(pn), Some(ct)) = (partname, content_type) {
                                map.add_override(pn, ct);
                            }
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(OpcError::XmlError(format!(
                        "Content types parse error: {}",
                        e
                    )));
                },
                _ => {},
            }
            buf.clear();
        }

        Ok(map)
    }

    /// Add a default content type mapping for a file extension.
    fn add_default(&mut self, extension: String, content_type: String) {
        self.defaults.insert(extension.to_lowercase(), content_type);
    }

    /// Add an override content type mapping for a specific partname.
    fn add_override(&mut self, partname: String, content_type: String) {
        self.overrides.insert(partname, content_type);
    }

    /// Get the content type for a partname.
    ///
    /// First checks for an override, then falls back to the default
    /// based on file extension.
    fn get(&self, pack_uri: &PackURI) -> Result<String> {
        // Check override first
        if let Some(ct) = self.overrides.get(pack_uri.as_str()) {
            return Ok(ct.clone());
        }

        // Fall back to default based on extension
        let ext = pack_uri.ext();
        if let Some(ct) = self.defaults.get(ext) {
            return Ok(ct.clone());
        }

        Err(OpcError::ContentTypeNotFound(pack_uri.to_string()))
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
    /// Open and parse an OPC package file.
    ///
    /// # Arguments
    /// * `phys_reader` - Physical package reader for accessing ZIP contents
    ///
    /// # Returns
    /// A new PackageReader with all parts and relationships loaded
    pub fn from_phys_reader<R: Read + Seek>(mut phys_reader: PhysPkgReader<R>) -> Result<Self> {
        // Parse content types
        let content_types_xml = phys_reader.content_types_xml()?;
        let content_types = ContentTypeMap::from_xml(&content_types_xml)?;

        // Get package-level relationships
        let package_uri = PackURI::new(PACKAGE_URI).map_err(OpcError::InvalidPackUri)?;
        let pkg_srels = Self::load_rels(&mut phys_reader, &package_uri)?;

        // Load all parts by walking the relationship graph
        let sparts = Self::load_parts(&mut phys_reader, &pkg_srels, &content_types)?;

        Ok(Self { pkg_srels, sparts })
    }

    /// Parse relationships from a .rels file.
    ///
    /// Uses efficient XML parsing with quick-xml to extract relationship information.
    fn load_rels<R: Read + Seek>(
        phys_reader: &mut PhysPkgReader<R>,
        source_uri: &PackURI,
    ) -> Result<SmallVec<[SerializedRelationship; 8]>> {
        let rels_xml = match phys_reader.rels_xml_for(source_uri)? {
            Some(xml) => xml,
            None => return Ok(SmallVec::new()), // No relationships file
        };

        let base_uri = source_uri.base_uri().to_string();
        let mut srels = SmallVec::new();
        let mut reader = Reader::from_reader(&rels_xml[..]);
        reader.config_mut().trim_text(true);

        let mut buf = Vec::new();

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Empty(ref e)) | Ok(Event::Start(ref e)) => {
                    if e.local_name().as_ref() == b"Relationship" {
                        let mut r_id = None;
                        let mut reltype = None;
                        let mut target_ref = None;
                        let mut target_mode = target_mode::INTERNAL.to_string();

                        for attr in e.attributes() {
                            let attr = attr?;
                            match attr.key.as_ref() {
                                b"Id" => r_id = Some(attr.unescape_value()?.to_string()),
                                b"Type" => reltype = Some(attr.unescape_value()?.to_string()),
                                b"Target" => target_ref = Some(attr.unescape_value()?.to_string()),
                                b"TargetMode" => target_mode = attr.unescape_value()?.to_string(),
                                _ => {},
                            }
                        }

                        if let (Some(id), Some(rt), Some(tr)) = (r_id, reltype, target_ref) {
                            srels.push(SerializedRelationship {
                                base_uri: base_uri.clone(),
                                r_id: id,
                                reltype: rt,
                                target_ref: tr,
                                target_mode,
                            });
                        }
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(OpcError::XmlError(format!("Rels parse error: {}", e))),
                _ => {},
            }
            buf.clear();
        }

        Ok(srels)
    }

    /// Load all parts by walking the relationship graph.
    ///
    /// Uses depth-first traversal to discover all parts reachable from
    /// package-level relationships. Tracks visited parts to avoid cycles.
    fn load_parts<R: Read + Seek>(
        phys_reader: &mut PhysPkgReader<R>,
        pkg_srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
    ) -> Result<Vec<SerializedPart>> {
        let mut sparts = Vec::new();
        let mut visited = HashMap::new();

        Self::walk_parts(
            phys_reader,
            pkg_srels,
            content_types,
            &mut visited,
            &mut sparts,
        )?;

        Ok(sparts)
    }

    /// Recursively walk the part relationship graph.
    ///
    /// This performs depth-first traversal of the part tree, loading each
    /// part and its relationships. Uses a visited map to detect cycles.
    fn walk_parts<R: Read + Seek>(
        phys_reader: &mut PhysPkgReader<R>,
        srels: &[SerializedRelationship],
        content_types: &ContentTypeMap,
        visited: &mut HashMap<String, ()>,
        sparts: &mut Vec<SerializedPart>,
    ) -> Result<()> {
        for srel in srels {
            // Skip external relationships
            if srel.is_external() {
                continue;
            }

            let partname = srel.target_partname()?;
            let partname_str = partname.to_string();

            // Skip if already visited
            if visited.contains_key(&partname_str) {
                continue;
            }
            visited.insert(partname_str, ());

            // Load part content
            let blob = phys_reader.blob_for(&partname)?;
            let content_type = content_types.get(&partname)?;

            // Load part relationships
            let part_srels = Self::load_rels(phys_reader, &partname)?;

            // Create serialized part
            let spart = SerializedPart {
                partname: partname.clone(),
                content_type,
                reltype: srel.reltype.clone(),
                blob,
                srels: part_srels.clone(),
            };
            sparts.push(spart);

            // Recursively load related parts
            Self::walk_parts(phys_reader, &part_srels, content_types, visited, sparts)?;
        }

        Ok(())
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

#[cfg(test)]
mod tests {
    use super::*;

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
}
