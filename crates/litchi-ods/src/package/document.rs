use litchi_core::{Error, Result};
use litchi_odf_common::core::{Content, OwnedPackage, Styles};
use std::{fs, path::Path};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.spreadsheet";

/// Validated ownership boundary for an ODS package.
pub struct SpreadsheetPackage {
    package: OwnedPackage,
    content: Content,
    styles: Option<Styles>,
}

impl SpreadsheetPackage {
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::from_bytes(fs::read(path)?)
    }

    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OwnedPackage::from_bytes(bytes)?;
        let mimetype = package.mimetype()?;
        if mimetype != MIMETYPE {
            return Err(Error::InvalidFormat(format!(
                "expected ODS package, found '{mimetype}'"
            )));
        }
        let content_bytes = package.get_file("content.xml")?;
        let content = Content::from_bytes(&content_bytes)?;
        let styles = package
            .has_file("styles.xml")?
            .then(|| package.get_file("styles.xml"))
            .transpose()?
            .map(|bytes| Styles::from_bytes(&bytes))
            .transpose()?;
        Ok(Self {
            package,
            content,
            styles,
        })
    }

    pub fn content_xml(&self) -> &str {
        self.content.xml_content()
    }

    pub fn styles_xml(&self) -> Option<&str> {
        self.styles.as_ref().map(Styles::xml_content)
    }

    pub fn package(&self) -> &OwnedPackage {
        &self.package
    }

    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_inner()
    }
}
