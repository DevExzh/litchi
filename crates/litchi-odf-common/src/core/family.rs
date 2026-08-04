//! Shared ownership for the simple packaged ODF families.

use super::{Content, Meta, OwnedPackage, Styles};
use litchi_core::{Error, Metadata, Result};
use std::{fs, path::Path};

const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;

/// A validated immutable ODF package with the standard XML parts decoded once.
///
/// Concrete family crates retain a small contextual wrapper around this type
/// so MIME and body validation remain visible at their package boundary while
/// archive, content, style, metadata, and file-list ownership stays shared.
pub struct FamilyPackage {
    archive: OwnedPackage,
    content: Content,
    styles: Option<Styles>,
    metadata: Option<Metadata>,
}

impl FamilyPackage {
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
        let archive = OwnedPackage::from_bytes(bytes)?;
        let found = archive.mimetype()?;
        if found != mimetype {
            return Err(Error::InvalidFormat(format!(
                "expected {family_name} package MIME type '{mimetype}', found '{found}'"
            )));
        }

        let content_bytes = archive.get_file("content.xml")?;
        if content_bytes.len() > MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{family_name} content.xml exceeds the family limit"
            )));
        }
        let content = Content::from_bytes(&content_bytes)?;
        if !content.xml_content().contains(body_marker) {
            return Err(Error::InvalidFormat(format!(
                "{family_name} content.xml has no expected body"
            )));
        }

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

    /// List all safe package paths.
    pub fn files(&self) -> Result<Vec<String>> {
        self.archive.files()
    }

    /// Consume the snapshot and return the original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.archive.into_inner()
    }
}
