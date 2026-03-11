//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.
//!
//! Uses soapberry-zip for high-performance zero-copy ZIP parsing.

use crate::common::{Error, Result};
use soapberry_zip::office::ArchiveReader;
use std::io::Read;

/// An ODF package (ZIP file containing XML documents)
///
/// Uses soapberry-zip for efficient lazy decompression.
pub struct Package<'data> {
    archive: ArchiveReader<'data>,
    #[allow(dead_code)]
    manifest: super::manifest::Manifest,
    mimetype: String,
}

/// Owned version of Package that owns the data buffer.
pub struct OwnedPackage {
    data: Vec<u8>,
}

#[allow(dead_code)]
impl OwnedPackage {
    /// Open an ODF package from a reader
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;

        // Validate the archive can be parsed
        let _ = ArchiveReader::new(&data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        Ok(Self { data })
    }

    /// Create an ODF package from bytes
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        // Validate the archive can be parsed
        let _ = ArchiveReader::new(&data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        Ok(Self { data })
    }

    /// Get a borrowed Package for accessing archive contents
    pub fn package(&self) -> Result<Package<'_>> {
        Package::new(&self.data)
    }

    /// Get the underlying data
    pub fn into_inner(self) -> Vec<u8> {
        self.data
    }

    /// Get a reference to the underlying data
    pub fn as_bytes(&self) -> &[u8] {
        &self.data
    }

    // Convenience methods that delegate to Package

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> Result<String> {
        let package = self.package()?;
        Ok(package.mimetype().to_string())
    }

    /// Get a file from the package by path
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let package = self.package()?;
        package.get_file(path)
    }

    /// Check if a file exists in the package
    pub fn has_file(&self, path: &str) -> Result<bool> {
        let package = self.package()?;
        Ok(package.has_file(path))
    }

    /// List all files in the package
    pub fn files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.files()
    }

    /// Get all embedded media files from the package.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.media_files()
    }
}

impl<'data> Package<'data> {
    /// Create a new Package from a byte slice
    pub fn new(data: &'data [u8]) -> Result<Self> {
        let archive = ArchiveReader::new(data)
            .map_err(|_| Error::InvalidFormat("Invalid ZIP archive".to_string()))?;

        // Read MIME type from mimetype file
        let mimetype = archive
            .read_string("mimetype")
            .map_err(|_| Error::InvalidFormat("No mimetype file found in ODF package".to_string()))?
            .trim()
            .to_string();

        // Parse the manifest
        let manifest = super::manifest::Manifest::from_archive_reader(&archive)?;

        Ok(Self {
            archive,
            manifest,
            mimetype,
        })
    }

    /// Get the MIME type from the mimetype file
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.archive
            .read(path)
            .map_err(|_| Error::InvalidFormat(format!("File not found: {}", path)))
    }

    /// Check if a file exists in the package
    pub fn has_file(&self, path: &str) -> bool {
        self.archive.contains(path)
    }

    /// Get the manifest
    #[allow(dead_code)]
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package
    pub fn files(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Get all embedded media files (images, etc.) from the package.
    ///
    /// This returns paths to all files in the Pictures/ directory and other media directories.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let all_files = self.files()?;
        Ok(all_files
            .into_iter()
            .filter(|path| {
                path.starts_with("Pictures/")
                    || path.starts_with("media/")
                    || path.starts_with("Object/")
                    || path.ends_with(".png")
                    || path.ends_with(".jpg")
                    || path.ends_with(".jpeg")
                    || path.ends_with(".gif")
                    || path.ends_with(".svg")
            })
            .collect())
    }

    /// Check if the package contains any media files.
    #[allow(dead_code)] // Reserved for future use
    pub fn has_media(&self) -> bool {
        self.media_files().map(|m| !m.is_empty()).unwrap_or(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    // Helper function to create a minimal ODF package (ZIP with mimetype and manifest)
    fn create_test_odf_package(mimetype: &str) -> Vec<u8> {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            // Write mimetype file (must be first and uncompressed for ODF)
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(mimetype.as_bytes()).unwrap();

            // Write manifest.xml
            let manifest_xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
    <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Pictures/image.png" manifest:media-type="image/png"/>
</manifest:manifest>"#;
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest_xml.as_bytes()).unwrap();

            // Write content.xml
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(b"<office:document-content/>").unwrap();

            // Write styles.xml
            zip.start_file("styles.xml", options).unwrap();
            zip.write_all(b"<office:document-styles/>").unwrap();

            // Write a picture
            zip.start_file("Pictures/image.png", options).unwrap();
            zip.write_all(b"PNG\x89\x50\x4e\x47\x0d\x0a\x1a\x0a")
                .unwrap();

            zip.finish().unwrap();
        }
        zip_buffer
    }

    fn create_test_ods_package() -> Vec<u8> {
        create_test_odf_package("application/vnd.oasis.opendocument.spreadsheet")
    }

    fn create_test_odp_package() -> Vec<u8> {
        create_test_odf_package("application/vnd.oasis.opendocument.presentation")
    }

    #[test]
    fn test_owned_package_from_bytes() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data);
        assert!(package.is_ok());
    }

    #[test]
    fn test_owned_package_from_reader() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let cursor = Cursor::new(data);
        let package = OwnedPackage::from_reader(cursor);
        assert!(package.is_ok());
    }

    #[test]
    fn test_owned_package_invalid_data() {
        let invalid_data = b"not a zip file".to_vec();
        let result = OwnedPackage::from_bytes(invalid_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_owned_package_into_inner() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data.clone()).unwrap();
        let inner = package.into_inner();
        assert!(!inner.is_empty());
    }

    #[test]
    fn test_owned_package_as_bytes() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data.clone()).unwrap();
        let bytes = package.as_bytes();
        assert!(!bytes.is_empty());
    }

    #[test]
    fn test_owned_package_mimetype() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_owned_package_mimetype_ods() {
        let data = create_test_ods_package();
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.spreadsheet"
        );
    }

    #[test]
    fn test_owned_package_mimetype_odp() {
        let data = create_test_odp_package();
        let package = OwnedPackage::from_bytes(data).unwrap();
        assert_eq!(
            package.mimetype().unwrap(),
            "application/vnd.oasis.opendocument.presentation"
        );
    }

    #[test]
    fn test_owned_package_get_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let content = package.get_file("content.xml");
        assert!(content.is_ok());
        assert_eq!(content.unwrap(), b"<office:document-content/>");
    }

    #[test]
    fn test_owned_package_get_file_not_found() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let result = package.get_file("nonexistent.xml");
        assert!(result.is_err());
    }

    #[test]
    fn test_owned_package_has_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        assert!(package.has_file("content.xml").unwrap());
        assert!(package.has_file("styles.xml").unwrap());
        assert!(!package.has_file("nonexistent.xml").unwrap());
    }

    #[test]
    fn test_owned_package_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let files = package.files().unwrap();
        assert!(files.contains(&"mimetype".to_string()));
        assert!(files.contains(&"content.xml".to_string()));
        assert!(files.contains(&"styles.xml".to_string()));
        assert!(files.contains(&"META-INF/manifest.xml".to_string()));
        assert!(files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_owned_package_media_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = OwnedPackage::from_bytes(data).unwrap();

        let media_files = package.media_files().unwrap();
        assert!(media_files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_package_new() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data);
        assert!(package.is_ok());
    }

    #[test]
    fn test_package_new_invalid_data() {
        let invalid_data = b"not a zip file";
        let result = Package::new(invalid_data);
        assert!(result.is_err());
    }

    #[test]
    fn test_package_mimetype() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();
        assert_eq!(
            package.mimetype(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_package_get_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let content = package.get_file("content.xml").unwrap();
        assert_eq!(content, b"<office:document-content/>");
    }

    #[test]
    fn test_package_get_file_not_found() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let result = package.get_file("nonexistent.xml");
        assert!(result.is_err());
    }

    #[test]
    fn test_package_has_file() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        assert!(package.has_file("content.xml"));
        assert!(!package.has_file("nonexistent.xml"));
    }

    #[test]
    fn test_package_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let files = package.files().unwrap();
        assert!(!files.is_empty());
        assert!(files.contains(&"content.xml".to_string()));
    }

    #[test]
    fn test_package_media_files() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let media_files = package.media_files().unwrap();
        assert!(media_files.contains(&"Pictures/image.png".to_string()));
    }

    #[test]
    fn test_package_has_media() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        assert!(package.has_media());
    }

    #[test]
    fn test_package_manifest() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data).unwrap();

        let manifest = package.manifest();
        assert_eq!(manifest.mimetype, "application/vnd.oasis.opendocument.text");
    }

    #[test]
    fn test_owned_package_package_method() {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let owned = OwnedPackage::from_bytes(data).unwrap();

        let package = owned.package();
        assert!(package.is_ok());
        assert_eq!(
            package.unwrap().mimetype(),
            "application/vnd.oasis.opendocument.text"
        );
    }

    #[test]
    fn test_package_media_files_various_formats() {
        use std::io::Write;

        let mut zip_buffer = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut zip_buffer));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);

            // Write mimetype file
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();

            // Write manifest.xml
            let manifest_xml = r#"<?xml version="1.0"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
    <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
</manifest:manifest>"#;
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest_xml.as_bytes()).unwrap();

            // Write various media files
            zip.start_file("Pictures/photo.jpg", options).unwrap();
            zip.write_all(b"fake jpg data").unwrap();

            zip.start_file("Pictures/chart.jpeg", options).unwrap();
            zip.write_all(b"fake jpeg data").unwrap();

            zip.start_file("media/animation.gif", options).unwrap();
            zip.write_all(b"fake gif data").unwrap();

            zip.start_file("Object/image.svg", options).unwrap();
            zip.write_all(b"<svg/>").unwrap();

            zip.start_file("media/diagram.png", options).unwrap();
            zip.write_all(b"fake png data").unwrap();

            zip.finish().unwrap();
        }

        let package = Package::new(&zip_buffer).unwrap();
        let media_files = package.media_files().unwrap();

        assert!(media_files.contains(&"Pictures/photo.jpg".to_string()));
        assert!(media_files.contains(&"Pictures/chart.jpeg".to_string()));
        assert!(media_files.contains(&"media/animation.gif".to_string()));
        assert!(media_files.contains(&"Object/image.svg".to_string()));
        assert!(media_files.contains(&"media/diagram.png".to_string()));
    }
}
