//! Owned UTF-8 storage shared by decoded ODF XML parts.

use litchi_core::{Error, Result};

/// A validated, immutable UTF-8 XML part.
#[allow(
    clippy::module_name_repetitions,
    reason = "`XmlPart` is the established public name for the XML storage primitive."
)]
#[derive(Debug)]
pub struct XmlPart {
    content: Box<str>,
}

impl XmlPart {
    /// Parse XML content from bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not valid UTF-8.
    pub fn from_bytes(bytes: &[u8]) -> Result<Self> {
        let content = String::from_utf8(bytes.to_vec())
            .map_err(|error| {
                Error::InvalidFormat(format!("Invalid UTF-8 in XML content: {error}"))
            })?
            .into_boxed_str();
        Ok(Self { content })
    }

    /// Borrow the raw XML content.
    #[must_use]
    pub fn content(&self) -> &str {
        &self.content
    }

    /// Borrow the raw XML bytes without another allocation.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.content.as_bytes()
    }
}
