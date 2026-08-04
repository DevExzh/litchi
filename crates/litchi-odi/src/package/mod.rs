//! Validated package ownership for this family.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{Content, Meta, OwnedPackage, Styles};
use std::{fs, path::Path};

pub const MIMETYPE: &str = "application/vnd.oasis.opendocument.image";

/// An immutable, validated package snapshot.
pub struct Package {
    archive: OwnedPackage,
    content: Content,
    styles: Option<Styles>,
    metadata: Option<Metadata>,
}

impl Package {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(fs::read(path)?)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let archive = OwnedPackage::from_bytes(bytes)?;
        let found = archive.mimetype()?;
        if found != MIMETYPE {
            return Err(Error::InvalidFormat(format!(
                "expected {MIMETYPE}, found '{found}'"
            )));
        }
        let bytes = archive.get_file("content.xml")?;
        let content = Content::from_bytes(&bytes)?;
        crate::codec::validate_content(content.xml_content())?;
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

    pub fn content_xml(&self) -> &str {
        self.content.xml_content()
    }
    pub fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }
    pub fn metadata(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }
    pub fn files(&self) -> Result<Vec<String>> {
        self.archive.files()
    }
    pub fn into_bytes(self) -> Vec<u8> {
        self.archive.into_inner()
    }
}
