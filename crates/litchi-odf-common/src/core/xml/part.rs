//! Owned UTF-8 storage shared by decoded ODF XML parts.

use litchi_core::{Error, Result};

/// A validated, immutable UTF-8 XML part.
#[derive(Debug)]
pub struct XmlPart {
    content: Box<str>,
}

impl XmlPart {
    /// Parse XML content from bytes.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let content = String::from_utf8(bytes.to_vec())
            .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in XML content".to_string()))?
            .into_boxed_str();
        Ok(Self { content })
    }

    /// Borrow the raw XML content.
    pub fn content(&self) -> &str {
        &self.content
    }

    /// Borrow the raw XML bytes without another allocation.
    #[allow(dead_code)]
    pub fn as_bytes(&self) -> &[u8] {
        self.content.as_bytes()
    }
}
