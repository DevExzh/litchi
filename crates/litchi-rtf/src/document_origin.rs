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
    pub fn from_rtf_value(value: i32) -> RtfResult<Self> {
        if value == 1 {
            Ok(Self::Version1)
        } else {
            Err(RtfError::MalformedDocument(
                "RTF fromhtml version must be 1".to_string(),
            ))
        }
    }
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
    pub fn effective_auto_format_type(self) -> DocumentAutoFormatType {
        self.auto_format_type.unwrap_or_default()
    }
    pub fn is_empty(self) -> bool {
        self.origin.is_none() && self.auto_format_type.is_none()
    }
}
