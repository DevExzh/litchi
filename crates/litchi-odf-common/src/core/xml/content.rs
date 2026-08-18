//! `content.xml` storage.

use super::XmlPart;
use litchi_core::Result;

/// Parsed `content.xml`.
#[derive(Debug)]
pub struct Content {
    xml: XmlPart,
}

impl Content {
    /// Parse `content.xml` from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not valid UTF-8.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            xml: XmlPart::from_bytes(bytes)?,
        })
    }

    /// Parse `content.xml` while consuming owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not valid UTF-8.
    pub fn from_vec(bytes: Vec<u8>) -> Result<Self> {
        Ok(Self {
            xml: XmlPart::from_vec(bytes)?,
        })
    }

    /// Borrow the raw `content.xml` text.
    #[must_use]
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }
}
