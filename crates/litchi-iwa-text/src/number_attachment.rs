//! Archive-free values for native page and slide-number text attachments.
//!
//! Native object lookup, protobuf conversion, text-storage resolution, and
//! package mutation stay in the owning IWA adapter. This module owns only the
//! checked values exchanged at that boundary.

use std::fmt;
use std::num::NonZeroU64;

use crate::position::TextPosition;

const MAX_ATTACHMENT_TEXT_BYTES: usize = 16 * 1_024;

/// Validation failures produced while constructing number-attachment values.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A native attachment object identifier was zero.
    ZeroObjectId,
    /// Attachment text exceeded the bounded semantic budget.
    TextTooLong {
        /// Number of UTF-8 bytes supplied by the caller.
        actual: usize,
        /// Maximum number of UTF-8 bytes accepted by the leaf.
        maximum: usize,
    },
    /// Attachment text contained a Unicode control character.
    TextControlCharacter,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::ZeroObjectId => {
                formatter.write_str("iWork number-attachment object identifier cannot be zero")
            },
            Self::TextTooLong { actual, maximum } => write!(
                formatter,
                "iWork number-attachment text is {actual} bytes; maximum is {maximum}"
            ),
            Self::TextControlCharacter => formatter
                .write_str("iWork number-attachment text cannot contain control characters"),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for number-attachment semantic values.
pub type Result<T> = std::result::Result<T, Error>;

/// Identifier of a native page or slide-number attachment.
//
// `NonZeroU64` keeps the invalid zero state out of the published value while
// retaining the native eight-byte representation.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentId(NonZeroU64);

impl TextNumberAttachmentId {
    /// Construct an identifier obtained from a native object reference.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroObjectId`] when `identifier` is zero.
    pub const fn from_object_id(identifier: u64) -> Result<Self> {
        match NonZeroU64::new(identifier) {
            Some(identifier) => Ok(Self(identifier)),
            None => Err(Error::ZeroObjectId),
        }
    }

    /// Construct an identifier from a native value without losing its bits.
    ///
    /// This is intentionally fallible: zero is not a valid native object
    /// identifier and is never representable by the semantic value.
    pub const fn from_native(identifier: u64) -> Result<Self> {
        Self::from_object_id(identifier)
    }

    /// Return the underlying package object identifier.
    #[must_use]
    pub const fn object_id(self) -> u64 {
        self.0.get()
    }
}

/// Native semantic role of a textual page or slide-number attachment.
//
// The raw discriminant is stored directly so future native values remain
// lossless in the same four-byte representation. Named associated constants
// provide the ergonomic known values without making an unknown enum variant
// large enough to widen the value.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TextNumberAttachmentKind(i32);

#[allow(
    non_upper_case_globals,
    reason = "PascalCase constants preserve the native semantic vocabulary."
)]
impl TextNumberAttachmentKind {
    /// The current page or slide number.
    pub const PageNumber: Self = Self(0);
    /// The total page or slide count.
    pub const PageCount: Self = Self(1);
    /// A footnote marker represented by the same native attachment family.
    pub const FootnoteMark: Self = Self(2);

    /// Decode a native kind discriminant without discarding unknown values.
    #[must_use]
    pub const fn from_raw(raw: i32) -> Self {
        Self(raw)
    }

    /// Decode a native kind discriminant without discarding unknown values.
    #[must_use]
    pub const fn from_native(raw: i32) -> Self {
        Self::from_raw(raw)
    }

    /// Return the lossless native kind discriminant.
    #[must_use]
    pub const fn as_raw(self) -> i32 {
        self.0
    }

    /// Return the lossless native kind discriminant.
    #[must_use]
    pub const fn native_value(self) -> i32 {
        self.as_raw()
    }

    /// Whether this value is not one of the currently known native kinds.
    #[must_use]
    pub const fn is_unsupported(self) -> bool {
        !matches!(self.0, 0..=2)
    }
}

/// Opaque native number-format selector.
//
// Apple stores this as an undocumented unsigned value. The newtype prevents
// it from being confused with object IDs, positions, or dimensions.
#[repr(transparent)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentFormat(u32);

impl TextNumberAttachmentFormat {
    /// Preserve an iWork number-format value without assigning invented
    /// semantics.
    #[must_use]
    pub const fn from_native_value(value: u32) -> Self {
        Self(value)
    }

    /// Preserve an iWork number-format value without assigning invented
    /// semantics.
    #[must_use]
    pub const fn from_native(value: u32) -> Self {
        Self::from_native_value(value)
    }

    /// Return the lossless native number-format value.
    #[must_use]
    pub const fn native_value(self) -> u32 {
        self.0
    }
}

/// Validated text stored inside a native number-attachment payload.
///
/// Empty text is valid because Keynote uses it as the string equivalent of a
/// slide-number attachment.
#[repr(transparent)]
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextNumberAttachmentText(Box<str>);

impl TextNumberAttachmentText {
    /// Maximum UTF-8 bytes retained by one attachment text value.
    pub const MAX_BYTES: usize = MAX_ATTACHMENT_TEXT_BYTES;

    /// Validate and own borrowed native attachment text.
    ///
    /// Validation occurs before allocation for borrowed input.
    pub fn new(value: &str) -> Result<Self> {
        validate_text(value)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain an existing boxed string without reallocating it.
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_text(&value)?;
        Ok(Self(value))
    }

    /// Borrow the text exactly as stored by iWork.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the text as an owned `String`.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for TextNumberAttachmentText {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for TextNumberAttachmentText {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<String> for TextNumberAttachmentText {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::from_boxed(value.into_boxed_str())
    }
}

impl TryFrom<Box<str>> for TextNumberAttachmentText {
    type Error = Error;

    fn try_from(value: Box<str>) -> Result<Self> {
        Self::from_boxed(value)
    }
}

/// Lossless writable payload of a native page or slide-number attachment.
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
    #[must_use]
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
    #[must_use]
    pub fn with_string_equivalent(mut self, value: TextNumberAttachmentText) -> Self {
        self.string_equivalent = Some(value);
        self
    }

    /// Replace the optional opaque number-format selector.
    #[must_use]
    pub const fn with_number_format(mut self, value: TextNumberAttachmentFormat) -> Self {
        self.number_format = Some(value);
        self
    }

    /// Replace the optional cached display value.
    #[must_use]
    pub fn with_string_value(mut self, value: TextNumberAttachmentText) -> Self {
        self.string_value = Some(value);
        self
    }

    /// Replace the optional native number-format name.
    #[must_use]
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
    /// Construct an attachment from already validated semantic values.
    #[must_use]
    pub const fn new(
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

fn validate_text(value: &str) -> Result<()> {
    if value.len() > MAX_ATTACHMENT_TEXT_BYTES {
        return Err(Error::TextTooLong {
            actual: value.len(),
            maximum: MAX_ATTACHMENT_TEXT_BYTES,
        });
    }
    if value.chars().any(char::is_control) {
        return Err(Error::TextControlCharacter);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    #[test]
    fn identifiers_are_nonzero_and_native_sized() {
        assert_eq!(size_of::<TextNumberAttachmentId>(), size_of::<u64>());
        assert_eq!(
            TextNumberAttachmentId::from_object_id(0),
            Err(Error::ZeroObjectId)
        );

        let identifier = TextNumberAttachmentId::from_native(42).unwrap_or_else(|error| {
            panic!("nonzero native identifiers should be accepted: {error}")
        });
        assert_eq!(identifier.object_id(), 42);
    }

    #[test]
    fn kinds_keep_known_and_unknown_discriminants_losslessly() {
        assert_eq!(size_of::<TextNumberAttachmentKind>(), size_of::<i32>());
        assert_eq!(TextNumberAttachmentKind::PageNumber.as_raw(), 0);
        assert_eq!(TextNumberAttachmentKind::PageCount.native_value(), 1);
        assert_eq!(TextNumberAttachmentKind::FootnoteMark.as_raw(), 2);

        for raw in [i32::MIN, -1, 3, 9_001, i32::MAX] {
            let kind = TextNumberAttachmentKind::from_native(raw);
            assert_eq!(kind.as_raw(), raw);
            assert!(kind.is_unsupported());
        }
    }

    #[test]
    fn text_is_bounded_and_validated_before_borrowed_allocation() {
        assert_eq!(TextNumberAttachmentText::new("").unwrap().as_str(), "");
        assert_eq!(
            TextNumberAttachmentText::new("plain text")
                .unwrap()
                .into_string(),
            "plain text"
        );
        assert_eq!(
            TextNumberAttachmentText::new("line\nbreak"),
            Err(Error::TextControlCharacter)
        );

        let oversized = "x".repeat(TextNumberAttachmentText::MAX_BYTES + 1);
        assert_eq!(
            TextNumberAttachmentText::new(&oversized),
            Err(Error::TextTooLong {
                actual: TextNumberAttachmentText::MAX_BYTES + 1,
                maximum: TextNumberAttachmentText::MAX_BYTES,
            })
        );
    }

    #[test]
    fn settings_builders_retain_each_optional_native_value() {
        let text = TextNumberAttachmentText::new("Page ").unwrap();
        let cached = TextNumberAttachmentText::new("7").unwrap();
        let name = TextNumberAttachmentText::new("decimal").unwrap();
        let settings = TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageNumber)
            .with_string_equivalent(text.clone())
            .with_number_format(TextNumberAttachmentFormat::from_native_value(u32::MAX))
            .with_string_value(cached.clone())
            .with_number_format_name(name.clone());

        assert_eq!(settings.kind, TextNumberAttachmentKind::PageNumber);
        assert_eq!(settings.string_equivalent.as_ref(), Some(&text));
        assert_eq!(
            settings
                .number_format
                .map(TextNumberAttachmentFormat::native_value),
            Some(u32::MAX)
        );
        assert_eq!(settings.string_value.as_ref(), Some(&cached));
        assert_eq!(settings.number_format_name.as_ref(), Some(&name));
    }

    #[test]
    fn attachment_is_composed_from_typed_compact_values() {
        assert_eq!(size_of::<TextNumberAttachmentFormat>(), size_of::<u32>());
        let attachment = TextNumberAttachment::new(
            TextNumberAttachmentId::from_object_id(7).unwrap(),
            TextPosition::from_utf16_code_units(4),
            TextNumberAttachmentSettings::new(TextNumberAttachmentKind::PageCount),
        );
        assert_eq!(attachment.id.object_id(), 7);
        assert_eq!(attachment.position.utf16_index(), 4);
        assert_eq!(
            attachment.settings.kind,
            TextNumberAttachmentKind::PageCount
        );
    }
}
