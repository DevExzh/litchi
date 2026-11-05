//! `meta.xml` storage and metadata projection.

use super::XmlPart;
use crate::core::XmlStreamLimits;
use litchi_core::Result;

/// Parsed `meta.xml`.
#[derive(Debug)]
pub struct Meta {
    xml: XmlPart,
    limits: XmlStreamLimits,
}

impl Meta {
    /// Copy `meta.xml` admitted under the raw policy's input-byte bound.
    ///
    /// # Errors
    ///
    /// Returns [`LimitKind::RawXml`](crate::LimitKind::RawXml) when the part exceeds the bound, and an invalid-format
    /// error when it is not UTF-8.
    pub fn from_bytes(
        bytes: &[u8],
        limits: XmlStreamLimits,
    ) -> std::result::Result<Self, crate::Error> {
        super::part::admit_raw(bytes.len(), limits)?;
        Ok(Self {
            xml: XmlPart::copy(bytes)?,
            limits,
        })
    }

    /// Adopt owned `meta.xml` bytes admitted under the raw policy's input-byte bound, without copying
    /// the buffer.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes`].
    pub fn from_vec(
        bytes: Vec<u8>,
        limits: XmlStreamLimits,
    ) -> std::result::Result<Self, crate::Error> {
        super::part::admit_raw(bytes.len(), limits)?;
        Ok(Self {
            xml: XmlPart::adopt(bytes)?,
            limits,
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
        crate::core::metadata::Metadata::from_xml(self.xml.content(), self.limits)
    }

    /// Extract common metadata while preserving parse failures.
    ///
    /// # Errors
    ///
    /// Returns an error when the metadata XML is malformed or invalid.
    pub fn try_extract_metadata(&self) -> Result<litchi_core::Metadata> {
        self.odf_metadata()?.try_into_core()
    }
}
