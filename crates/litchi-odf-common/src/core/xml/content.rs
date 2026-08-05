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
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            xml: XmlPart::from_bytes(bytes)?,
        })
    }

    /// Borrow the raw `content.xml` text.
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }
}
