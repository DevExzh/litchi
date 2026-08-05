//! `styles.xml` storage.

use super::XmlPart;
use litchi_core::Result;

/// Parsed `styles.xml`.
#[derive(Debug)]
pub struct Styles {
    xml: XmlPart,
}

impl Styles {
    /// Parse `styles.xml` from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        Ok(Self {
            xml: XmlPart::from_bytes(bytes)?,
        })
    }

    /// Borrow the raw `styles.xml` text.
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }
}
