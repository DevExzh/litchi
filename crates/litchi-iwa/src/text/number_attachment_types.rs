//! Strict public value types for native iWork page/slide-number attachments.

use crate::{Error, Result};

use litchi_iwa_text::position::TextPosition;

const MAX_ATTACHMENT_TEXT_BYTES: usize = 16 * 1_024;

/// Identifier of a native number-attachment object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentId(u64);

impl TextNumberAttachmentId {
    /// Construct an identifier obtained from a previously read attachment.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork number-attachment object identifier cannot be zero".to_owned(),
            ));
        }
        Ok(Self(identifier))
    }

    /// Return the underlying package object identifier.
    pub const fn object_id(self) -> u64 {
        self.0
    }

    pub(crate) const fn from_native(identifier: u64) -> Self {
        Self(identifier)
    }
}

/// Native semantic role of a textual number attachment.
///
/// Keynote uses `PageNumber` for the current slide number. Unknown raw values
/// remain writable so newer iWork variants round-trip without normalization.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum TextNumberAttachmentKind {
    PageNumber,
    PageCount,
    FootnoteMark,
    Unknown(i32),
}

impl TextNumberAttachmentKind {
    pub(crate) const fn from_raw(raw: i32) -> Self {
        match raw {
            0 => Self::PageNumber,
            1 => Self::PageCount,
            2 => Self::FootnoteMark,
            unknown => Self::Unknown(unknown),
        }
    }

    pub(crate) const fn as_raw(self) -> i32 {
        match self {
            Self::PageNumber => 0,
            Self::PageCount => 1,
            Self::FootnoteMark => 2,
            Self::Unknown(raw) => raw,
        }
    }
}

/// Opaque native number-format selector.
///
/// Apple stores this as an undocumented unsigned value. The newtype prevents
/// it from being confused with object IDs, positions, or dimensions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentFormat(u32);

impl TextNumberAttachmentFormat {
    /// Preserve an iWork number-format value without assigning invented semantics.
    pub const fn from_native_value(value: u32) -> Self {
        Self(value)
    }

    /// Return the lossless native number-format value.
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Validated text stored inside a native number-attachment payload.
///
/// Empty strings are accepted because Keynote uses one as the string
/// equivalent of its slide-number attachment.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentText(Box<str>);

impl TextNumberAttachmentText {
    /// Validate and construct lossless native attachment text.
    pub fn new(value: impl Into<Box<str>>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_ATTACHMENT_TEXT_BYTES {
            return Err(Error::ParseError(format!(
                "iWork number-attachment text exceeds {MAX_ATTACHMENT_TEXT_BYTES} bytes"
            )));
        }
        if value.chars().any(char::is_control) {
            return Err(Error::ParseError(
                "iWork number-attachment text cannot contain control characters".to_owned(),
            ));
        }
        Ok(Self(value))
    }

    /// Return the text exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextNumberAttachmentText {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Lossless writable payload of a native number attachment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextNumberAttachmentSettings {
    /// Required semantic role of the attachment.
    pub kind: TextNumberAttachmentKind,
    /// Optional accessibility/localization equivalent.
    pub string_equivalent: Option<TextNumberAttachmentText>,
    /// Optional opaque native number-format selector.
    pub number_format: Option<TextNumberAttachmentFormat>,
    /// Optional cached display value.
    pub string_value: Option<TextNumberAttachmentText>,
    /// Optional native number-format name.
    pub number_format_name: Option<TextNumberAttachmentText>,
}

impl TextNumberAttachmentSettings {
    /// Construct a minimal native attachment of the requested kind.
    pub const fn new(kind: TextNumberAttachmentKind) -> Self {
        Self {
            kind,
            string_equivalent: None,
            number_format: None,
            string_value: None,
            number_format_name: None,
        }
    }

    /// Replace the optional string equivalent.
    pub fn with_string_equivalent(mut self, value: TextNumberAttachmentText) -> Self {
        self.string_equivalent = Some(value);
        self
    }

    /// Replace the optional opaque number-format selector.
    pub const fn with_number_format(mut self, value: TextNumberAttachmentFormat) -> Self {
        self.number_format = Some(value);
        self
    }

    /// Replace the optional cached display value.
    pub fn with_string_value(mut self, value: TextNumberAttachmentText) -> Self {
        self.string_value = Some(value);
        self
    }

    /// Replace the optional native number-format name.
    pub fn with_number_format_name(mut self, value: TextNumberAttachmentText) -> Self {
        self.number_format_name = Some(value);
        self
    }
}

/// One native number attachment at a U+FFFC text position.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextNumberAttachment {
    /// Native attachment object identifier.
    pub id: TextNumberAttachmentId,
    /// UTF-16 position of the object-replacement character.
    pub position: TextPosition,
    /// Losslessly decoded native payload.
    pub settings: TextNumberAttachmentSettings,
}

impl TextNumberAttachment {
    pub(crate) const fn new(
        id: TextNumberAttachmentId,
        position: TextPosition,
        settings: TextNumberAttachmentSettings,
    ) -> Self {
        Self {
            id,
            position,
            settings,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn public_values_are_strict_and_lossless() {
        assert!(TextNumberAttachmentId::from_object_id(0).is_err());
        assert!(TextNumberAttachmentText::new("line\nbreak").is_err());
        assert!(TextNumberAttachmentText::new("").is_ok());
        for raw in [-8, 0, 1, 2, 19] {
            assert_eq!(TextNumberAttachmentKind::from_raw(raw).as_raw(), raw);
        }
        let format = TextNumberAttachmentFormat::from_native_value(u32::MAX);
        assert_eq!(format.native_value(), u32::MAX);
    }
}
