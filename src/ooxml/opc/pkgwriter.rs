//! Package writer for OPC packages.
//!
//! This module provides functionality to serialize and write OPC packages to disk,
//! including writing the [Content_Types].xml, relationships, and all parts.

use crate::ooxml::opc::constants::content_type as ct;
use crate::ooxml::opc::error::Result;
use crate::ooxml::opc::package::OpcPackage;
use crate::ooxml::opc::packuri::{CONTENT_TYPES_URI, PACKAGE_URI, PackURI};
use crate::ooxml::opc::phys_pkg::PhysPkgWriter;
use std::collections::HashMap;
use std::path::Path;

/// Package writer that serializes an OPC package to a ZIP file.
///
/// This is the main entry point for saving packages. It handles writing:
/// - [Content_Types].xml
/// - _rels/.rels (package relationships)
/// - All parts and their relationships
///
/// # Example
///
/// ```no_run
/// use litchi::ooxml::opc::package::OpcPackage;
/// use litchi::ooxml::opc::pkgwriter::PackageWriter;
///
/// let mut pkg = OpcPackage::new();
/// // ... add parts to package ...
/// PackageWriter::write("output.docx", &pkg)?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
pub struct PackageWriter;

impl PackageWriter {
    /// Write an OPC package to a file.
    ///
    /// # Arguments
    /// * `path` - Path where the package should be written
    /// * `package` - The OPC package to write
    pub fn write<P: AsRef<Path>>(path: P, package: &OpcPackage) -> Result<()> {
        let bytes = Self::to_bytes(package)?;
        std::fs::write(path, bytes)?;
        Ok(())
    }

    /// Write an OPC package to a stream.
    ///
    /// # Arguments
    /// * `writer` - A writer that implements Write
    /// * `package` - The OPC package to write
    pub fn write_to_stream<W: std::io::Write>(mut writer: W, package: &OpcPackage) -> Result<()> {
        let bytes = Self::to_bytes(package)?;
        writer.write_all(&bytes)?;
        Ok(())
    }

    /// Serialize an OPC package to bytes.
    ///
    /// # Arguments
    /// * `package` - The OPC package to serialize
    ///
    /// # Returns
    /// The serialized package as a byte vector
    pub fn to_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
        let mut phys_writer = PhysPkgWriter::new();

        // Write [Content_Types].xml
        Self::write_content_types(&mut phys_writer, package)?;

        // Write package-level relationships (_rels/.rels)
        Self::write_pkg_rels(&mut phys_writer, package)?;

        // Write all parts and their relationships
        Self::write_parts(&mut phys_writer, package)?;

        // Finish writing and return the bytes
        phys_writer.finish()
    }

    /// Write the [Content_Types].xml part.
    ///
    /// This file maps file extensions and part names to content types.
    fn write_content_types(phys_writer: &mut PhysPkgWriter, package: &OpcPackage) -> Result<()> {
        let cti = ContentTypesItem::from_package(package);
        let blob = cti.to_xml();

        let content_types_uri = PackURI::new(CONTENT_TYPES_URI)
            .map_err(crate::ooxml::opc::error::OpcError::InvalidPackUri)?;
        phys_writer.write(&content_types_uri, blob.as_bytes())?;

        Ok(())
    }

    /// Write package-level relationships.
    fn write_pkg_rels(phys_writer: &mut PhysPkgWriter, package: &OpcPackage) -> Result<()> {
        let package_uri = PackURI::new(PACKAGE_URI)
            .map_err(crate::ooxml::opc::error::OpcError::InvalidPackUri)?;
        let rels_uri = package_uri
            .rels_uri()
            .map_err(crate::ooxml::opc::error::OpcError::InvalidPackUri)?;
        let rels_xml = package.rels().to_xml();
        phys_writer.write(&rels_uri, rels_xml.as_bytes())?;

        Ok(())
    }

    /// Write all parts and their relationships.
    fn write_parts(phys_writer: &mut PhysPkgWriter, package: &OpcPackage) -> Result<()> {
        for part in package.iter_parts() {
            // Write the part itself
            let blob = part.blob();
            phys_writer.write(part.partname(), blob)?;

            // Write the part's relationships if it has any
            if !part.rels().is_empty() {
                let rels_uri = part
                    .partname()
                    .rels_uri()
                    .map_err(crate::ooxml::opc::error::OpcError::InvalidPackUri)?;
                let rels_xml = part.rels().to_xml();
                phys_writer.write(&rels_uri, rels_xml.as_bytes())?;
            }
        }

        Ok(())
    }
}

/// Helper for building [Content_Types].xml content.
///
/// Manages Default and Override elements for content type mapping.
struct ContentTypesItem {
    /// Default content types by extension
    defaults: HashMap<String, String>,

    /// Override content types by partname
    overrides: HashMap<String, String>,
}

impl ContentTypesItem {
    /// Create a new ContentTypesItem.
    fn new() -> Self {
        let mut defaults = HashMap::new();

        // Add standard defaults
        defaults.insert("rels".to_string(), ct::OPC_RELATIONSHIPS.to_string());
        defaults.insert("xml".to_string(), ct::XML.to_string());

        Self {
            defaults,
            overrides: HashMap::new(),
        }
    }

    /// Build ContentTypesItem from an OPC package.
    fn from_package(package: &OpcPackage) -> Self {
        let mut cti = Self::new();

        for part in package.iter_parts() {
            cti.add_content_type(part.partname(), part.content_type());
        }

        cti
    }

    /// Add a content type for a part.
    ///
    /// Uses a default mapping if the extension matches a well-known type,
    /// otherwise uses an override for the specific partname.
    fn add_content_type(&mut self, partname: &PackURI, content_type: &str) {
        let ext = partname.ext();

        // Check if this is a standard default mapping
        if Self::is_default_content_type(ext, content_type) {
            self.defaults
                .insert(ext.to_string(), content_type.to_string());
        } else {
            self.overrides
                .insert(partname.to_string(), content_type.to_string());
        }
    }

    /// Check if an extension/content-type pair is a standard default.
    fn is_default_content_type(ext: &str, content_type: &str) -> bool {
        matches!(
            (ext, content_type),
            ("rels", ct::OPC_RELATIONSHIPS)
                | ("xml", ct::XML)
                | ("bin", ct::XLSB_BIN)
                | ("png", "image/png")
                | ("jpg", "image/jpeg")
                | ("jpeg", "image/jpeg")
                | ("gif", "image/gif")
                | ("emf", "image/x-emf")
                | ("wmf", "image/x-wmf")
        )
    }

    /// Generate the XML for [Content_Types].xml.
    fn to_xml(&self) -> String {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push('\n');
        xml.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        xml.push('\n');

        // Write Default elements (sorted by extension)
        let mut exts: Vec<_> = self.defaults.keys().collect();
        exts.sort();
        for ext in exts {
            let content_type = &self.defaults[ext];
            xml.push_str(&format!(
                r#"  <Default Extension="{}" ContentType="{}"/>"#,
                Self::escape_xml(ext),
                Self::escape_xml(content_type)
            ));
            xml.push('\n');
        }

        // Write Override elements (sorted by partname)
        let mut partnames: Vec<_> = self.overrides.keys().collect();
        partnames.sort();
        for partname in partnames {
            let content_type = &self.overrides[partname];
            xml.push_str(&format!(
                r#"  <Override PartName="{}" ContentType="{}"/>"#,
                Self::escape_xml(partname),
                Self::escape_xml(content_type)
            ));
            xml.push('\n');
        }

        xml.push_str("</Types>");

        xml
    }

    /// Escape XML special characters.
    #[inline]
    fn escape_xml(s: &str) -> String {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
            .replace('\'', "&apos;")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_content_types_xml() {
        let mut cti = ContentTypesItem::new();
        cti.defaults
            .insert("png".to_string(), "image/png".to_string());
        cti.overrides.insert(
            "/word/document.xml".to_string(),
            ct::WML_DOCUMENT_MAIN.to_string(),
        );

        let xml = cti.to_xml();

        assert!(xml.contains(r#"<Default Extension="png" ContentType="image/png"/>"#));
        assert!(xml.contains(r#"<Override PartName="/word/document.xml""#));
    }

    #[test]
    fn test_xml_escaping() {
        let escaped = ContentTypesItem::escape_xml(r#"<foo & "bar">"#);
        assert_eq!(escaped, "&lt;foo &amp; &quot;bar&quot;&gt;");
    }
}
