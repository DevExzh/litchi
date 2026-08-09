use crate::{RtfError, RtfResult};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DocumentOrigin {
    PlainTextEmail,
    HtmlEmail { version: Option<HtmlEmailVersion> },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HtmlEmailVersion {
    Version1,
}

impl HtmlEmailVersion {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_rtf_value(value: i32) -> RtfResult<Self> {
        if value == 1 {
            Ok(Self::Version1)
        } else {
            Err(RtfError::MalformedDocument(
                "RTF fromhtml version must be 1".to_string(),
            ))
        }
    }
    #[must_use]
    pub fn rtf_value(self) -> i32 {
        1
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum DocumentAutoFormatType {
    #[default]
    General,
    Letter,
    Email,
}

impl DocumentAutoFormatType {
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_rtf_value(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::General),
            1 => Ok(Self::Letter),
            2 => Ok(Self::Email),
            _ => Err(RtfError::MalformedDocument(
                "RTF document type must be in the range 0 through 2".to_string(),
            )),
        }
    }
    #[must_use]
    pub fn rtf_value(self) -> i32 {
        match self {
            Self::General => 0,
            Self::Letter => 1,
            Self::Email => 2,
        }
    }
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentOriginMetadata {
    pub origin: Option<DocumentOrigin>,
    pub auto_format_type: Option<DocumentAutoFormatType>,
}

impl DocumentOriginMetadata {
    #[must_use]
    pub fn effective_auto_format_type(self) -> DocumentAutoFormatType {
        self.auto_format_type.unwrap_or_default()
    }
    #[must_use]
    pub fn is_empty(self) -> bool {
        self.origin.is_none() && self.auto_format_type.is_none()
    }
}
