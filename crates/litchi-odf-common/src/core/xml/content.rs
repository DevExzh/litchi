//! `content.xml` storage.

use super::XmlPart;
use crate::limits::{SemanticLimitKind, SemanticLimits};

/// Parsed `content.xml`.
#[derive(Debug)]
pub struct Content {
    xml: XmlPart,
}

impl Content {
    /// Copy `content.xml` admitted under the caller's content-bytes bound.
    ///
    /// # Errors
    ///
    /// Returns [`SemanticLimitKind::ContentBytes`] when the part exceeds the bound, and an invalid-format
    /// error when it is not UTF-8.
    pub fn from_bytes(bytes: &[u8], limits: SemanticLimits) -> Result<Self, crate::Error> {
        limits.check_usize(SemanticLimitKind::ContentBytes, bytes.len())?;
        Ok(Self {
            xml: XmlPart::copy(bytes)?,
        })
    }

    /// Adopt owned `content.xml` bytes admitted under the caller's content-bytes bound, without copying
    /// the buffer.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes`].
    pub fn from_vec(bytes: Vec<u8>, limits: SemanticLimits) -> Result<Self, crate::Error> {
        limits.check_usize(SemanticLimitKind::ContentBytes, bytes.len())?;
        Ok(Self {
            xml: XmlPart::adopt(bytes)?,
        })
    }

    /// Borrow the raw `content.xml` text.
    #[must_use]
    pub fn xml_content(&self) -> &str {
        self.xml.content()
    }
}
