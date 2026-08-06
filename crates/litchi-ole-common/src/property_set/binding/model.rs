//! Semantic property-set binding values.

use super::super::model::{
    DOCUMENT_SUMMARY_INFORMATION_FMTID, Guid, SUMMARY_INFORMATION_FMTID,
    USER_DEFINED_PROPERTIES_FMTID,
};
use litchi_cfb::OleError;
use std::fmt;

/// FMTID of the shared `GlobalInfo` property-set binding.
pub const GLOBAL_INFO_FMTID: Guid = Guid::from_bytes([
    0x00, 0x6F, 0x61, 0x56, 0x54, 0xC1, 0xCE, 0x11, 0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B,
]);

/// FMTID of the shared `ImageContents` property-set binding.
pub const IMAGE_CONTENTS_FMTID: Guid = Guid::from_bytes([
    0x00, 0x64, 0x61, 0x56, 0x54, 0xC1, 0xCE, 0x11, 0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B,
]);

/// FMTID of the shared `ImageInfo` property-set binding.
pub const IMAGE_INFO_FMTID: Guid = Guid::from_bytes([
    0x00, 0x65, 0x61, 0x56, 0x54, 0xC1, 0xCE, 0x11, 0x85, 0x53, 0x00, 0xAA, 0x00, 0xA1, 0xF9, 0x5B,
]);

/// A standard or GUID-derived OLE Property Set binding.
///
/// `UserDefinedProperties` intentionally has its own FMTID even though the
/// standard name is the same as `DocumentSummaryInformation`.  Parsing that
/// shared name resolves to the canonical document-summary binding because a
/// name alone cannot distinguish the two FMTIDs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Binding {
    /// The `\x05SummaryInformation` binding.
    SummaryInformation,
    /// The `\x05DocumentSummaryInformation` binding.
    DocumentSummaryInformation,
    /// The user-defined section in the `\x05DocumentSummaryInformation` stream.
    UserDefinedProperties,
    /// The `\x05GlobalInfo` binding.
    GlobalInfo,
    /// The `\x05ImageContents` binding.
    ImageContents,
    /// The `\x05ImageInfo` binding.
    ImageInfo,
    /// A non-special FMTID encoded as a 26-character standard binding name.
    Custom(Guid),
}

impl Binding {
    /// Resolve a FMTID to its typed standard binding.
    pub fn from_format_identifier(format_identifier: Guid) -> Self {
        match format_identifier {
            SUMMARY_INFORMATION_FMTID => Self::SummaryInformation,
            DOCUMENT_SUMMARY_INFORMATION_FMTID => Self::DocumentSummaryInformation,
            USER_DEFINED_PROPERTIES_FMTID => Self::UserDefinedProperties,
            GLOBAL_INFO_FMTID => Self::GlobalInfo,
            IMAGE_CONTENTS_FMTID => Self::ImageContents,
            IMAGE_INFO_FMTID => Self::ImageInfo,
            value => Self::Custom(value),
        }
    }

    /// Create a custom binding without allocating or inspecting its FMTID.
    pub const fn custom(format_identifier: Guid) -> Self {
        Self::Custom(format_identifier)
    }

    /// Return the FMTID represented by this binding.
    pub const fn format_identifier(self) -> Guid {
        match self {
            Self::SummaryInformation => SUMMARY_INFORMATION_FMTID,
            Self::DocumentSummaryInformation => DOCUMENT_SUMMARY_INFORMATION_FMTID,
            Self::UserDefinedProperties => USER_DEFINED_PROPERTIES_FMTID,
            Self::GlobalInfo => GLOBAL_INFO_FMTID,
            Self::ImageContents => IMAGE_CONTENTS_FMTID,
            Self::ImageInfo => IMAGE_INFO_FMTID,
            Self::Custom(value) => value,
        }
    }

    /// Return the canonical CFB binding name without allocating.
    pub fn name(self) -> BindingName {
        super::codec::encode(self.format_identifier())
    }

    /// Parse a standard CFB binding name into its typed FMTID.
    pub fn from_name(name: &str) -> Result<Self, OleError> {
        super::codec::decode(name).map(Self::from_format_identifier)
    }

    /// Whether the binding's section is stored in the document-summary stream.
    pub const fn uses_document_summary_stream(self) -> bool {
        matches!(
            self,
            Self::DocumentSummaryInformation | Self::UserDefinedProperties
        )
    }
}

/// A validated standard CFB property-set name.
///
/// The value is always exactly 27 bytes: the control-character prefix plus
/// the maximum 26-character GUID-derived suffix.  Short special names are
/// zero-filled internally but [`as_str`](Self::as_str) exposes only their
/// actual canonical length.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct BindingName {
    bytes: [u8; super::validation::MAX_NAME_BYTES],
    len: u8,
}

impl BindingName {
    /// Return the encoded CFB path component.
    pub fn as_str(&self) -> &str {
        std::str::from_utf8(&self.bytes[..usize::from(self.len)])
            .expect("BindingName stores only valid UTF-8 bytes")
    }

    /// Return the encoded CFB path component as bytes.
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes[..usize::from(self.len)]
    }

    /// Return the number of bytes in the path component.
    pub const fn len(self) -> usize {
        self.len as usize
    }

    /// Whether this is the empty binding name.  Valid names are never empty.
    pub const fn is_empty(self) -> bool {
        self.len == 0
    }

    pub(super) fn from_bytes(bytes: [u8; super::validation::MAX_NAME_BYTES], len: usize) -> Self {
        Self {
            bytes,
            len: len as u8,
        }
    }
}

impl AsRef<str> for BindingName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl fmt::Debug for BindingName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_tuple("BindingName")
            .field(&self.as_str())
            .finish()
    }
}

impl fmt::Display for BindingName {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.as_str())
    }
}
