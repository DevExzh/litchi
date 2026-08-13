//! ODF package (ZIP archive) handling functionality.
//!
//! This module provides utilities for working with ODF files as ZIP archives,
//! including reading files, checking existence, and basic package operations.
//!
//! Uses soapberry-zip for high-performance zero-copy ZIP parsing.

use crate::package::{self, Archive, PreparedArchive};
use litchi_core::{Error, Result};
use std::io::Read;
use std::sync::Arc;
use zeroize::Zeroizing;

/// An ODF package (ZIP file containing XML documents).
///
/// Uses soapberry-zip for efficient lazy decompression.
#[allow(
    clippy::module_name_repetitions,
    reason = "`Package` is the established public ODF archive reader name."
)]
pub struct Package<'data> {
    archive: Archive<'data>,
    manifest: super::manifest::Manifest,
    mimetype: String,
    password: Option<&'data str>,
}

/// Owned version of [`Package`] that owns the data buffer.
#[derive(Clone)]
#[allow(
    clippy::module_name_repetitions,
    reason = "`OwnedPackage` distinguishes the owning public archive handle."
)]
pub struct OwnedPackage {
    data: Arc<Vec<u8>>,
    index: PreparedArchive,
    password: Option<Zeroizing<String>>,
}

impl OwnedPackage {
    /// Open an ODF package from a reader.
    ///
    /// # Errors
    ///
    /// Returns an error when reading the input or parsing the ZIP archive
    /// fails.
    pub fn from_reader<R: Read>(mut reader: R) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes(data)
    }

    /// Create an ODF package from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_bytes(data: Vec<u8>) -> Result<Self> {
        Self::from_shared_bytes(Arc::new(data))
    }

    /// Adopt shared ODF package bytes without copying the archive buffer.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_shared_bytes(data: Arc<Vec<u8>>) -> Result<Self> {
        Self::from_shared_bytes_with_policy(
            data,
            soapberry_zip::office::ArchiveValidationPolicy::Normalized,
        )
    }

    pub(crate) fn from_prepared_bytes(data: Vec<u8>) -> Result<Self> {
        Self::from_shared_bytes_with_policy(
            Arc::new(data),
            soapberry_zip::office::ArchiveValidationPolicy::StrictPackage,
        )
    }

    fn from_shared_bytes_with_policy(
        data: Arc<Vec<u8>>,
        policy: soapberry_zip::office::ArchiveValidationPolicy,
    ) -> Result<Self> {
        #[cfg(test)]
        package::note_index_build();
        let index = Arc::new(
            soapberry_zip::office::IndexedArchive::from_reader_with_limits_and_policy(
                Arc::clone(&data),
                u64::try_from(data.len()).map_err(|_| {
                    Error::InvalidFormat("ODF package length exceeds ZIP reader limits".to_string())
                })?,
                soapberry_zip::office::ArchiveLimits::default(),
                policy,
            )
            .map_err(|error| Error::InvalidFormat(format!("Invalid ZIP archive: {error}")))?,
        );
        Ok(Self {
            data,
            index,
            password: None,
        })
    }

    /// Open an ODF package and retain a password for lazy entry decryption.
    ///
    /// # Errors
    ///
    /// Returns an error when reading the input or parsing the ZIP archive
    /// fails.
    pub fn from_reader_with_password<R: Read>(
        mut reader: R,
        password: impl Into<String>,
    ) -> Result<Self> {
        let mut data = Vec::new();
        reader.read_to_end(&mut data)?;
        Self::from_bytes_with_password(data, password)
    }

    /// Open ODF bytes and retain a password for lazy entry decryption.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes do not form a valid ZIP archive.
    pub fn from_bytes_with_password(data: Vec<u8>, password: impl Into<String>) -> Result<Self> {
        let mut package = Self::from_shared_bytes(Arc::new(data))?;
        package.password = Some(Zeroizing::new(password.into()));
        Ok(package)
    }

    /// Get a borrowed [`Package`] for accessing archive contents.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive MIME type or manifest cannot be
    /// decoded.
    pub fn package(&self) -> Result<Package<'_>> {
        Package::new_with_prepared(
            Arc::clone(&self.index),
            self.password.as_ref().map(|password| password.as_str()),
        )
    }

    /// Get the underlying data.
    #[must_use]
    pub fn into_inner(self) -> Vec<u8> {
        let Self {
            data,
            index,
            password: _,
        } = self;
        match Arc::try_unwrap(index) {
            Ok(index) => {
                drop(data);
                match Arc::try_unwrap(index.into_zip_archive().into_inner()) {
                    Ok(data) => data,
                    Err(data) => (*data).clone(),
                }
            },
            Err(index) => {
                drop(index);
                Arc::try_unwrap(data).unwrap_or_else(|data| (*data).clone())
            },
        }
    }

    /// Get a reference to the underlying data.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.data.as_slice()
    }

    /// Clone the shared handle to the exact archive allocation.
    #[must_use]
    pub fn shared_bytes(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.data)
    }

    /// Return a stable identity for the retained archive index.
    #[doc(hidden)]
    #[must_use]
    pub fn prepared_index_identity(&self) -> usize {
        Arc::as_ptr(&self.index) as usize
    }

    // Convenience methods that delegate to Package

    /// Get the MIME type from the mimetype file.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn mimetype(&self) -> Result<String> {
        let package = self.package()?;
        Ok(package.mimetype().to_string())
    }

    /// Get a file from the package by path.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata, requested entry, or entry
    /// decryption is invalid.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let package = self.package()?;
        package.get_file(path)
    }

    /// Check if a file exists in the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn has_file(&self, path: &str) -> Result<bool> {
        let package = self.package()?;
        Ok(package.has_file(path))
    }

    /// Check whether a package member uses ZIP Store compression.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        self.index
            .is_stored(path)
            .map_err(|error| Error::InvalidFormat(error.to_string()))
    }

    /// List all files in the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.files()
    }

    /// Get all embedded media files from the package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package metadata cannot be decoded.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let package = self.package()?;
        package.media_files()
    }

    /// Read inert document and macro signature metadata from the package.
    ///
    /// This does not verify cryptographic signatures or execute macro content.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded.
    pub fn digital_signatures(&self) -> Result<crate::signature::DigitalSignatures> {
        self.package()?.digital_signatures()
    }

    /// Cryptographically verify document signatures without making any PKI trust claim.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded or verified.
    pub fn verify_document_signatures(
        &self,
    ) -> Result<Vec<crate::signature::SignatureVerification>> {
        crate::signature::verify_package(&self.data)
    }
}

impl<'data> Package<'data> {
    /// Create a new [`Package`] from a byte slice.
    ///
    /// # Errors
    ///
    /// Returns an error when the archive MIME type or manifest cannot be
    /// decoded.
    pub fn new(data: &'data [u8]) -> Result<Self> {
        Self::new_with_password(data, None)
    }

    fn new_with_password(data: &'data [u8], password: Option<&'data str>) -> Result<Self> {
        let archive = Archive::new(data)?;

        Self::from_archive(archive, password)
    }

    fn new_with_prepared(index: PreparedArchive, password: Option<&'data str>) -> Result<Self> {
        let archive = Archive::from_prepared(index);

        Self::from_archive(archive, password)
    }

    fn from_archive(archive: Archive<'data>, password: Option<&'data str>) -> Result<Self> {
        // Read MIME type from mimetype file
        let mimetype = archive
            .read_string("mimetype")
            .map_err(|error| {
                Error::InvalidFormat(format!("No mimetype file found in ODF package: {error}"))
            })?
            .trim()
            .to_string();

        // Parse the manifest
        let manifest = super::manifest::Manifest::parse(&archive.read_manifest_xml()?)?;

        Ok(Self {
            archive,
            manifest,
            mimetype,
            password,
        })
    }

    /// Get the MIME type from the mimetype file.
    #[must_use]
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Get a file from the package by path.
    ///
    /// # Errors
    ///
    /// Returns an error when the entry does not exist, does not meet encrypted
    /// package requirements, or cannot be decrypted.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        let bytes = self
            .archive
            .read(path)
            .map_err(|error| Error::InvalidFormat(format!("File not found: {path}: {error}")))?;
        let Some(entry) = self.manifest.get_entry(path) else {
            return Ok(bytes);
        };
        let Some(encryption) = &entry.encryption else {
            return Ok(bytes);
        };
        if !self.archive.is_stored(path).map_err(|error| {
            Error::InvalidFormat(format!(
                "Unable to inspect encrypted ODF entry '{path}': {error}"
            ))
        })? {
            return Err(Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' must use ZIP Store"
            )));
        }
        let password = self.password.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Password required for encrypted ODF entry '{path}'"
            ))
        })?;
        let size = entry.size.ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Encrypted ODF entry '{path}' has no plaintext size"
            ))
        })?;
        super::encryption::decrypt_entry(&bytes, password, encryption, size)
    }

    /// Check if a file exists in the package.
    #[must_use]
    pub fn has_file(&self, path: &str) -> bool {
        self.archive.contains(path)
    }

    /// Check whether a package member uses ZIP Store compression.
    pub fn is_stored(&self, path: &str) -> Result<bool> {
        self.archive.is_stored(path)
    }

    /// Get the manifest.
    #[must_use]
    pub fn manifest(&self) -> &super::manifest::Manifest {
        &self.manifest
    }

    /// List all files in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn files(&self) -> Result<Vec<String>> {
        Ok(self.archive.file_names().map(String::from).collect())
    }

    /// Get all embedded media files (images, etc.) from the package.
    ///
    /// This returns paths to all files in the Pictures/ directory and other media directories.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn media_files(&self) -> Result<Vec<String>> {
        let all_files = self.files()?;
        Ok(all_files
            .into_iter()
            .filter(|path| package::is_media_path(path))
            .collect())
    }

    /// Check if the package contains any media files.
    ///
    /// # Errors
    ///
    /// Returns an error if the archive cannot enumerate its entries.
    pub fn has_media(&self) -> Result<bool> {
        Ok(!self.media_files()?.is_empty())
    }

    /// Read inert document and macro signature metadata from the package.
    ///
    /// This does not verify cryptographic signatures or execute macro content.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or its signature metadata cannot be
    /// decoded.
    pub fn digital_signatures(&self) -> Result<crate::signature::DigitalSignatures> {
        use crate::signature::{
            DOCUMENT_SIGNATURE_PATH, MACRO_SIGNATURE_PATH, parse_signature_container,
        };

        let mut result = crate::signature::DigitalSignatures::default();
        if self.has_file(DOCUMENT_SIGNATURE_PATH) {
            result.document_signatures =
                parse_signature_container(&self.get_file(DOCUMENT_SIGNATURE_PATH)?)?;
        }
        if self.has_file(MACRO_SIGNATURE_PATH) {
            result.macro_signatures =
                parse_signature_container(&self.get_file(MACRO_SIGNATURE_PATH)?)?;
        }
        Ok(result)
    }
}

impl package::PackageLookup for Package<'_> {
    fn has_file(&self, path: &str) -> bool {
        self.has_file(path)
    }

    fn media_type(&self, path: &str) -> Option<&str> {
        self.manifest.get_media_type(path)
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible ZIP setup operations so assertions can focus on package behavior."
)]
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
    fn test_package_has_media() -> Result<()> {
        let data = create_test_odf_package("application/vnd.oasis.opendocument.text");
        let package = Package::new(&data)?;

        assert!(package.has_media()?);
        Ok(())
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
