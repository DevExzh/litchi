//! Shared ownership for the simple packaged ODF families.

use super::{Content, Meta, OwnedPackage, Styles};
use litchi_core::{Error, Metadata, Result};
use std::{fs, path::Path};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// Validate the bounded, family-specific `content.xml` contract shared by
/// package facades and detached family builders.
pub fn validate_content_part(xml: &str, body_marker: &str, family_name: &str) -> Result<()> {
    if xml.len() > MAX_CONTENT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml exceeds the family limit"
        )));
    }
    if !xml.contains(body_marker) {
        return Err(Error::InvalidFormat(format!(
            "{family_name} content.xml has no expected body"
        )));
    }
    Ok(())
}

/// A validated immutable ODF package with the standard XML parts decoded once.
///
/// Concrete family crates retain a small contextual wrapper around this type
/// so MIME and body validation remain visible at their package boundary while
/// archive, content, style, metadata, and file-list ownership stays shared.
pub struct Package {
    archive: OwnedPackage,
    content: Content,
    styles: Option<Styles>,
    metadata: Option<Metadata>,
}

impl Package {
    /// Open a package after validating its MIME type and content root marker.
    pub fn open(
        path: impl AsRef<Path>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_bytes(fs::read(path)?, mimetype, body_marker, family_name)
    }

    /// Decode a package after validating its MIME type and content root marker.
    pub fn from_bytes(
        bytes: Vec<u8>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_bytes(bytes)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Decode a password-encrypted package after validating its MIME type and
    /// content root marker.
    pub fn from_bytes_with_password(
        bytes: Vec<u8>,
        password: impl Into<String>,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        Self::from_owned_package(
            OwnedPackage::from_bytes_with_password(bytes, password)?,
            mimetype,
            body_marker,
            family_name,
        )
    }

    /// Adopt an already parsed archive without reparsing its ZIP structure.
    pub fn from_owned_package(
        archive: OwnedPackage,
        mimetype: &str,
        body_marker: &str,
        family_name: &str,
    ) -> Result<Self> {
        let found = archive.mimetype()?;
        if found != mimetype {
            return Err(Error::InvalidFormat(format!(
                "expected {family_name} package MIME type '{mimetype}', found '{found}'"
            )));
        }

        let content_bytes = archive.get_file("content.xml")?;
        let content_xml = std::str::from_utf8(&content_bytes)
            .map_err(|_| Error::InvalidFormat(format!("{family_name} content.xml is not UTF-8")))?;
        validate_content_part(content_xml, body_marker, family_name)?;
        let content = Content::from_bytes(&content_bytes)?;

        let styles = archive
            .has_file("styles.xml")?
            .then(|| archive.get_file("styles.xml"))
            .transpose()?
            .map(|bytes| Styles::from_bytes(&bytes))
            .transpose()?;
        let metadata = archive
            .has_file("meta.xml")?
            .then(|| archive.get_file("meta.xml"))
            .transpose()?
            .map(|bytes| Meta::from_bytes(&bytes))
            .transpose()?
            .map(|meta| meta.try_extract_metadata())
            .transpose()?;

        Ok(Self {
            archive,
            content,
            styles,
            metadata,
        })
    }

    /// Return the decoded content XML.
    pub fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    /// Return the optional decoded styles XML.
    pub fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    /// Return the optional common metadata snapshot.
    pub fn metadata(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }

    /// Borrow the owned package for family-specific package edits.
    pub fn package(&self) -> &OwnedPackage {
        &self.archive
    }

    /// Borrow the original archive bytes without allocating.
    pub fn as_bytes(&self) -> &[u8] {
        self.archive.as_bytes()
    }

    /// List all safe package paths.
    pub fn files(&self) -> Result<Vec<String>> {
        self.archive.files()
    }

    /// Consume the snapshot and return the original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.archive.into_inner()
    }
}

#[cfg(test)]
mod tests {
    use super::{OwnedPackage, Package, validate_content_part};
    use std::io::{Cursor, Write};

    const MIMETYPE: &str = "application/vnd.oasis.opendocument.presentation";
    const CONTENT: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:presentation/></office:body></office:document-content>"#;

    fn package() -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        let mut zip = zip::ZipWriter::new(&mut output);
        let options = zip::write::SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored);
        zip.start_file("mimetype", options).unwrap();
        zip.write_all(MIMETYPE.as_bytes()).unwrap();
        zip.start_file("META-INF/manifest.xml", options).unwrap();
        zip.write_all(
            format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="{MIMETYPE}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/><m:file-entry m:full-path="styles.xml" m:media-type="text/xml"/></m:manifest>"#
            )
            .as_bytes(),
        )
        .unwrap();
        zip.start_file("content.xml", options).unwrap();
        zip.write_all(CONTENT.as_bytes()).unwrap();
        zip.start_file("styles.xml", options).unwrap();
        zip.write_all(b"<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>")
            .unwrap();
        zip.finish().unwrap();
        output.into_inner()
    }

    #[test]
    fn validates_and_reuses_shared_package_state() {
        let bytes = package();
        let value =
            Package::from_bytes(bytes.clone(), MIMETYPE, "<office:presentation", "ODP").unwrap();
        assert_eq!(value.content_xml(), CONTENT);
        assert!(value.styles_xml().is_some());
        assert_eq!(value.as_bytes(), bytes.as_slice());
        assert!(value.files().unwrap().contains(&"content.xml".to_string()));

        let owned = OwnedPackage::from_bytes(bytes.clone()).unwrap();
        let adopted =
            Package::from_owned_package(owned, MIMETYPE, "<office:presentation", "ODP").unwrap();
        assert_eq!(adopted.as_bytes(), bytes.as_slice());
    }

    #[test]
    fn rejects_wrong_mime_and_body_marker() {
        let bytes = package();
        assert!(
            Package::from_bytes(bytes.clone(), "text/plain", "<office:presentation", "ODP")
                .is_err()
        );
        assert!(Package::from_bytes(bytes, MIMETYPE, "<office:text", "ODP").is_err());
    }

    #[test]
    fn validates_detached_content_with_the_same_family_contract() {
        assert!(validate_content_part("<office:drawing/>", "<office:drawing", "ODG").is_ok());
        let error = validate_content_part("<office:text/>", "<office:drawing", "ODG")
            .unwrap_err()
            .to_string();
        assert!(error.contains("ODG content.xml has no expected body"));
    }
}
