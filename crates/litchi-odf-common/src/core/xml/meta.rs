//! `meta.xml` storage and metadata projection.

use super::XmlPart;
use litchi_core::Result;

/// Parsed `meta.xml`.
#[derive(Debug)]
pub struct Meta {
    xml: XmlPart,
}

impl Meta {
    /// Parse `meta.xml` from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not valid UTF-8.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            xml: XmlPart::from_bytes(bytes)?,
        })
    }

    /// Borrow the raw `meta.xml` text.
    #[must_use]
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }

    /// Parse the complete `OpenDocument` metadata model.
    ///
    /// # Errors
    ///
    /// Returns an error when the metadata XML is malformed or invalid.
    pub fn odf_metadata(&self) -> Result<crate::core::metadata::Metadata> {
        crate::core::metadata::Metadata::from_xml(self.xml.content())
    }

    /// Extract common metadata while preserving parse failures.
    ///
    /// # Errors
    ///
    /// Returns an error when the metadata XML is malformed or invalid.
    pub fn try_extract_metadata(&self) -> Result<litchi_core::Metadata> {
        Ok(self.odf_metadata()?.into())
    }
}
