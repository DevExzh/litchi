/// Objects that implement reading and writing OPC packages.
///
/// This module provides the main OpcPackage type, which represents an Open Packaging
/// Convention package in memory. It manages parts, relationships, and provides
/// high-level operations for working with office documents.

use std::collections::HashMap;
use std::io::{Read, Seek};
use std::path::Path;
use crate::ooxml::opc::constants::relationship_type;
use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::{PackURI, PACKAGE_URI};
use crate::ooxml::opc::part::{Part, PartFactory};
use crate::ooxml::opc::phys_pkg::PhysPkgReader;
use crate::ooxml::opc::pkgreader::PackageReader;
use crate::ooxml::opc::rel::Relationships;

/// Main API class for working with OPC packages.
///
/// OpcPackage represents an Open Packaging Convention package in memory,
/// providing access to parts, relationships, and package-level operations.
/// Uses efficient data structures and minimal cloning for best performance.
pub struct OpcPackage {
    /// Package-level relationships
    rels: Relationships,
    
    /// All parts in the package, indexed by partname
    /// Using Box<dyn Part> for trait objects to allow different part types
    parts: HashMap<String, Box<dyn Part>>,
}

impl OpcPackage {
    /// Create a new empty OPC package.
    pub fn new() -> Self {
        Self {
            rels: Relationships::new(PACKAGE_URI.to_string()),
            parts: HashMap::new(),
        }
    }

    /// Open an OPC package from a file.
    ///
    /// # Arguments
    /// * `path` - Path to the package file (.docx, .xlsx, .pptx, etc.)
    ///
    /// # Returns
    /// A new OpcPackage instance loaded with the package contents
    ///
    /// # Example
    /// ```no_run
    /// use litchi::ooxml::opc::package::OpcPackage;
    ///
    /// let pkg = OpcPackage::open("document.docx").unwrap();
    /// ```
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self> {
        let phys_reader = PhysPkgReader::open(path)?;
        Self::from_phys_reader(phys_reader)
    }

    /// Load an OPC package from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read + Seek
    pub fn from_reader<R: Read + Seek>(reader: R) -> Result<Self> {
        let phys_reader = PhysPkgReader::new(reader)?;
        Self::from_phys_reader(phys_reader)
    }

    /// Load an OPC package from a physical package reader.
    fn from_phys_reader<R: Read + Seek>(phys_reader: PhysPkgReader<R>) -> Result<Self> {
        let pkg_reader = PackageReader::from_phys_reader(phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Unmarshal a package from a package reader.
    ///
    /// This is the main deserialization logic that converts serialized parts
    /// and relationships into the in-memory object graph.
    fn unmarshal(pkg_reader: PackageReader) -> Result<Self> {
        let mut package = Self::new();

        // First pass: Create all parts
        let mut parts_map: HashMap<String, Box<dyn Part>> = HashMap::new();
        
        for spart in pkg_reader.iter_sparts() {
            let part = PartFactory::load(
                spart.partname.clone(),
                spart.content_type.clone(),
                spart.blob.clone(), // TODO: Optimize to avoid clone
            )?;
            parts_map.insert(spart.partname.to_string(), part);
        }

        // Second pass: Load package relationships
        for srel in pkg_reader.pkg_srels() {
            package.rels.add_relationship(
                srel.reltype.clone(),
                srel.target_ref.clone(),
                srel.r_id.clone(),
                srel.is_external(),
            );
        }
        
        // Load part relationships
        for spart in pkg_reader.iter_sparts() {
            if let Some(part) = parts_map.get_mut(&spart.partname.to_string()) {
                for srel in &spart.srels {
                    part.rels_mut().add_relationship(
                        srel.reltype.clone(),
                        srel.target_ref.clone(),
                        srel.r_id.clone(),
                        srel.is_external(),
                    );
                }
            }
        }

        package.parts = parts_map;
        Ok(package)
    }

    /// Get a reference to the main document part.
    ///
    /// For Word documents, this is the document.xml part.
    /// For Excel, the workbook.xml part.
    /// For PowerPoint, the presentation.xml part.
    pub fn main_document_part(&self) -> Result<&dyn Part> {
        let rel = self.rels.part_with_reltype(relationship_type::OFFICE_DOCUMENT)?;
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Get a part by its partname.
    ///
    /// # Arguments
    /// * `partname` - The PackURI of the part to retrieve
    pub fn get_part(&self, partname: &PackURI) -> Result<&dyn Part> {
        self.parts
            .get(partname.as_str())
            .map(|b| &**b as &dyn Part)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Get a mutable reference to a part by its partname.
    pub fn get_part_mut(&mut self, partname: &PackURI) -> Result<&mut dyn Part> {
        self.parts
            .get_mut(partname.as_str())
            .map(|b| &mut **b as &mut dyn Part)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Get a part by relationship type from the package level.
    ///
    /// # Arguments
    /// * `reltype` - The relationship type URI
    pub fn part_by_reltype(&self, reltype: &str) -> Result<&dyn Part> {
        let rel = self.rels.part_with_reltype(reltype)?;
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Add a new part to the package.
    ///
    /// # Arguments
    /// * `part` - The part to add
    pub fn add_part(&mut self, part: Box<dyn Part>) {
        let partname = part.partname().to_string();
        self.parts.insert(partname, part);
    }

    /// Get an iterator over all parts in the package.
    pub fn iter_parts(&self) -> impl Iterator<Item = &dyn Part> {
        self.parts.values().map(|b| &**b as &dyn Part)
    }

    /// Get the number of parts in the package.
    pub fn part_count(&self) -> usize {
        self.parts.len()
    }

    /// Get a reference to the package-level relationships.
    pub fn rels(&self) -> &Relationships {
        &self.rels
    }

    /// Get a mutable reference to the package-level relationships.
    pub fn rels_mut(&mut self) -> &mut Relationships {
        &mut self.rels
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
        let rel = self.rels.get_or_add(reltype, partname);
        rel.r_id().to_string()
    }

    /// Find the next available partname for a part template.
    ///
    /// Useful for creating new parts with sequential numbering (e.g., image1.png, image2.png).
    ///
    /// # Arguments
    /// * `template` - A format string with a %d placeholder for the number
    ///
    /// # Example
    /// ```no_run
    /// # use litchi::ooxml::opc::package::OpcPackage;
    /// # let mut pkg = OpcPackage::new();
    /// let next_image = pkg.next_partname("/word/media/image%d.png");
    /// ```
    pub fn next_partname(&self, template: &str) -> Result<PackURI> {
        let mut n = 1u32;
        loop {
            let candidate = template.replace("%d", &n.to_string());
            if !self.parts.contains_key(&candidate) {
                return PackURI::new(candidate).map_err(OpcError::InvalidPackUri);
            }
            n += 1;
            if n > 10000 {
                // Safety limit to prevent infinite loops
                return Err(OpcError::InvalidPackUri(
                    "Too many parts, cannot find next partname".to_string()
                ));
            }
        }
    }

    /// Check if a part exists in the package.
    pub fn contains_part(&self, partname: &PackURI) -> bool {
        self.parts.contains_key(partname.as_str())
    }
}

impl Default for OpcPackage {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn create_minimal_docx() -> Vec<u8> {
        let mut zip_data = Vec::new();
        {
            let cursor = Cursor::new(&mut zip_data);
            let mut writer = ZipWriter::new(cursor);
            let options = SimpleFileOptions::default();

            // Add [Content_Types].xml
            writer.start_file("[Content_Types].xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"#).unwrap();

            // Add _rels/.rels
            writer.start_file("_rels/.rels", options).unwrap();
            writer.write_all(br#"<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"#).unwrap();

            // Add word/document.xml
            writer.start_file("word/document.xml", options).unwrap();
            writer.write_all(br#"<?xml version="1.0"?>
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <body><p><t>Test</t></p></body>
</document>"#).unwrap();

            writer.finish().unwrap();
        }
        zip_data
    }

    #[test]
    fn test_open_package() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor).unwrap();

        assert!(pkg.part_count() > 0);
    }

    #[test]
    fn test_main_document_part() {
        let zip_data = create_minimal_docx();
        let cursor = Cursor::new(zip_data);
        let pkg = OpcPackage::from_reader(cursor).unwrap();

        let main_part = pkg.main_document_part().unwrap();
        assert_eq!(main_part.content_type(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml");
    }
}

