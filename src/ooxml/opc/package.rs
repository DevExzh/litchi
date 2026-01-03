//! Objects that implement reading and writing OPC packages.
//!
//! This module provides the main OpcPackage type, which represents an Open Packaging
//! Convention package in memory. It manages parts, relationships, and provides
//! high-level operations for working with office documents.

use crate::ooxml::opc::constants::relationship_type;
use crate::ooxml::opc::error::{OpcError, Result};
use crate::ooxml::opc::packuri::{PACKAGE_URI, PackURI};
use crate::ooxml::opc::part::{Part, PartFactory};
use crate::ooxml::opc::phys_pkg::{OwnedPhysPkgReader, PhysPkgReader};
use crate::ooxml::opc::pkgreader::PackageReader;
use crate::ooxml::opc::rel::Relationships;
use std::collections::HashMap;
use std::io::Read;
use std::path::Path;

/// Main API class for working with OPC packages.
///
/// OpcPackage represents an Open Packaging Convention package in memory,
/// providing access to parts, relationships, and package-level operations.
/// Uses efficient data structures and minimal cloning for best performance.
pub struct OpcPackage {
    /// Package-level relationships
    rels: Relationships,

    /// All parts in the package, indexed by partname
    /// Using Box<dyn Part + Send + Sync> for trait objects to allow different part types
    /// PackURI keys avoid string allocations compared to String keys
    parts: HashMap<PackURI, Box<dyn Part + Send + Sync>>,
}

impl std::fmt::Debug for OpcPackage {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("OpcPackage")
            .field("rels", &self.rels)
            .field("parts_count", &self.parts.len())
            .finish()
    }
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
        let owned_reader = OwnedPhysPkgReader::open(path)?;
        let phys_reader = owned_reader.reader()?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Load an OPC package from a reader.
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        let owned_reader = OwnedPhysPkgReader::from_reader(reader)?;
        let phys_reader = owned_reader.reader()?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Load an OPC package from a byte slice.
    ///
    /// # Arguments
    /// * `data` - The ZIP archive data as a byte slice
    pub fn from_bytes(data: &[u8]) -> Result<Self> {
        let phys_reader = PhysPkgReader::new(data)?;
        let pkg_reader = PackageReader::from_phys_reader(&phys_reader)?;
        Self::unmarshal(pkg_reader)
    }

    /// Unmarshal a package from a package reader.
    ///
    /// This is the main deserialization logic that converts serialized parts
    /// and relationships into the in-memory object graph.
    ///
    /// Optimized to minimize clones by consuming the package reader and moving data.
    fn unmarshal(mut pkg_reader: PackageReader) -> Result<Self> {
        let mut package = Self::new();

        // Get ownership of package relationships and parts
        let pkg_srels = pkg_reader.take_pkg_srels();
        let sparts = pkg_reader.take_sparts();

        // Pre-allocate with known capacity to avoid reallocations
        let mut parts_map: HashMap<PackURI, Box<dyn Part + Send + Sync>> =
            HashMap::with_capacity(sparts.len());

        // Create all parts - move data instead of cloning
        for spart in sparts {
            let partname = spart.partname.clone(); // Need to clone partname for the HashMap key
            let mut part = PartFactory::load(
                spart.partname,     // Move
                spart.content_type, // Move
                spart.blob,         // Move (blob is Arc internally if large)
            )?;

            // Load part relationships
            for srel in spart.srels {
                let is_external = srel.is_external(); // Evaluate before move
                part.rels_mut().add_relationship(
                    srel.reltype,    // Move
                    srel.target_ref, // Move
                    srel.r_id,       // Move
                    is_external,
                );
            }

            parts_map.insert(partname, part);
        }

        // Load package relationships - move instead of clone
        for srel in pkg_srels {
            let is_external = srel.is_external(); // Evaluate before move
            package.rels.add_relationship(
                srel.reltype,    // Move
                srel.target_ref, // Move
                srel.r_id,       // Move
                is_external,
            );
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
        let rel = self
            .rels
            .part_with_reltype(relationship_type::OFFICE_DOCUMENT)?;
        let partname = rel.target_partname()?;
        self.get_part(&partname)
    }

    /// Get a part by its partname.
    ///
    /// # Arguments
    /// * `partname` - The PackURI of the part to retrieve
    pub fn get_part(&self, partname: &PackURI) -> Result<&dyn Part> {
        self.parts
            .get(partname)
            .map(|b| &**b as &dyn Part)
            .ok_or_else(|| OpcError::PartNotFound(partname.to_string()))
    }

    /// Get a mutable reference to a part by its partname.
    pub fn get_part_mut(&mut self, partname: &PackURI) -> Result<&mut dyn Part> {
        self.parts
            .get_mut(partname)
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
    pub fn add_part(&mut self, part: Box<dyn Part + Send + Sync>) {
        let partname = part.partname().clone();
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

    /// Add an external relationship (e.g., for hyperlinks).
    ///
    /// # Arguments
    /// * `target_url` - External URL target
    /// * `reltype` - Relationship type
    ///
    /// # Returns
    /// The relationship ID (e.g., "rId1")
    pub fn relate_to_external(&mut self, target_url: &str, reltype: &str) -> String {
        self.rels.get_or_add_ext_rel(reltype, target_url)
    }

    /// Get mutable access to package-level relationships.
    ///
    /// Useful for advanced relationship management.
    pub fn relationships_mut(&mut self) -> &mut Relationships {
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
    /// # use litchi::ooxml::opc::package::OpcPackage;
    /// # let mut pkg = OpcPackage::new();
    /// let next_image = pkg.next_partname("/word/media/image%d.png");
    /// ```
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
            let candidate_str = std::str::from_utf8(&candidate_bytes)
                .map_err(|_| OpcError::InvalidPackUri("Invalid UTF-8 in partname".to_string()))?;

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
    pub fn contains_part(&self, partname: &PackURI) -> bool {
        self.parts.contains_key(partname)
    }

    /// Save the package to a file.
    ///
    /// Writes the complete OPC package including all parts, relationships,
    /// and content types to a ZIP file.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    ///
    /// # Example
    /// ```no_run
    /// use litchi::ooxml::opc::package::OpcPackage;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// pkg.save("output.docx")?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn save<P: AsRef<Path>>(&self, path: P) -> Result<()> {
        crate::ooxml::opc::pkgwriter::PackageWriter::write(path, self)
    }

    /// Save the package to a stream.
    ///
    /// Writes the complete OPC package including all parts, relationships,
    /// and content types to a writer stream.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write + Seek
    ///
    /// # Example
    /// ```no_run
    /// use litchi::ooxml::opc::package::OpcPackage;
    /// use std::fs::File;
    ///
    /// let mut pkg = OpcPackage::new();
    /// // ... add parts to package ...
    /// let file = File::create("output.docx")?;
    /// pkg.to_stream(file)?;
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn to_stream<W: std::io::Write + std::io::Seek>(&self, writer: W) -> Result<()> {
        crate::ooxml::opc::pkgwriter::PackageWriter::write_to_stream(writer, self)
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
    use soapberry_zip::office::StreamingArchiveWriter;
    use std::io::Cursor;

    fn create_minimal_docx() -> Vec<u8> {
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
        writer
            .write_deflated(
                "word/document.xml",
                br#"<?xml version="1.0"?>
<document xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <body><p><t>Test</t></p></body>
</document>"#,
            )
            .unwrap();

        writer.finish_to_bytes().unwrap()
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
        assert_eq!(
            main_part.content_type(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        );
    }
}
