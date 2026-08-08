//! Archive-free semantic values for native iWork text hyperlinks.
//!
//! This module owns only validated hyperlink identity, target text, and the
//! UTF-16 range covered by a link. Native object lookup, protobuf fields, and
//! wire mutation remain in the concrete IWA adapter.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Opaque IDs keep their raw adapter module adjacent to private representations."
)]

use std::num::NonZeroU64;

use crate::position::TextRange;

/// Maximum UTF-8 byte length accepted for a hyperlink target.
pub const MAX_TARGET_BYTES: usize = 8 * 1_024;

/// Validation failures produced while constructing a text hyperlink value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The hyperlink identity is zero.
    ZeroId,
    /// The hyperlink target is empty.
    EmptyTarget,
    /// The hyperlink target exceeds [`MAX_TARGET_BYTES`].
    TargetTooLong,
    /// The hyperlink target has whitespace at either boundary.
    TargetSurroundingWhitespace,
    /// The hyperlink target contains a Unicode control character.
    TargetControlCharacter,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let message = match self {
            Self::ZeroId => "text hyperlink identifier must be non-zero",
            Self::EmptyTarget => "text hyperlink target must not be empty",
            Self::TargetTooLong => "text hyperlink target exceeds 8192 UTF-8 bytes",
            Self::TargetSurroundingWhitespace => {
                "text hyperlink target must not have leading or trailing whitespace"
            },
            Self::TargetControlCharacter => {
                "text hyperlink target must not contain control characters"
            },
        };
        formatter.write_str(message)
    }
}

impl std::error::Error for Error {}

/// Result type for hyperlink value construction.
pub type Result<T> = std::result::Result<T, Error>;

/// A compact, validated semantic hyperlink identity.
#[allow(
    clippy::module_name_repetitions,
    reason = "TextHyperlinkId is the established public name of this semantic value."
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct TextHyperlinkId(NonZeroU64);

/// Explicit native-boundary conversions for the opaque hyperlink handle.
#[allow(
    clippy::arbitrary_source_item_ordering,
    reason = "Keep the explicit raw adapter boundary adjacent to its opaque handle."
)]
pub mod raw {
    use super::{Result, TextHyperlinkId};

    /// Validate a native hyperlink object identifier at the IWA boundary.
    ///
    /// # Errors
    ///
    /// Returns [`super::Error::ZeroId`] when `identifier` is zero.
    pub const fn from_object_id(identifier: u64) -> Result<TextHyperlinkId> {
        TextHyperlinkId::from_raw(identifier)
    }

    /// Recover a native hyperlink object identifier inside an adapter.
    #[must_use]
    pub const fn object_id(identifier: TextHyperlinkId) -> u64 {
        identifier.into_raw()
    }
}

impl TextHyperlinkId {
    const fn from_raw(identifier: u64) -> Result<Self> {
        match NonZeroU64::new(identifier) {
            Some(non_zero) => Ok(Self(non_zero)),
            None => Err(Error::ZeroId),
        }
    }

    const fn into_raw(self) -> u64 {
        self.0.get()
    }
}

/// A validated, opaque native hyperlink target.
///
/// Targets are deliberately not parsed into URL schemes. Ordinary URLs,
/// `mailto:` values, Keynote targets such as `?slide=next`, and future native
/// forms are stored and returned byte-for-byte as supplied.
#[allow(
    clippy::module_name_repetitions,
    reason = "TextHyperlinkTarget is the established public name of this semantic value."
)]
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextHyperlinkTarget(Box<str>);

impl TextHyperlinkTarget {
    /// Maximum UTF-8 byte length accepted for one target.
    pub const MAX_BYTES: usize = MAX_TARGET_BYTES;

    /// Validate and own a borrowed hyperlink target.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyTarget`] when the target is empty,
    /// [`Error::TargetTooLong`] when it exceeds the bound,
    /// [`Error::TargetSurroundingWhitespace`] when it has surrounding
    /// whitespace, or [`Error::TargetControlCharacter`] when it contains a
    /// Unicode control character.
    pub fn new(value: &str) -> Result<Self> {
        validate_target(value)?;
        Ok(Self(value.into()))
    }

    /// Validate and retain a boxed hyperlink target without reallocating it.
    ///
    /// # Errors
    ///
    /// Returns the same typed failures as [`Self::new`].
    pub fn from_boxed(value: Box<str>) -> Result<Self> {
        validate_target(&value)?;
        Ok(Self(value))
    }

    /// Borrow the target exactly as stored by iWork.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the target without allocating.
    #[must_use]
    pub fn into_boxed_str(self) -> Box<str> {
        self.0
    }
}

impl AsRef<str> for TextHyperlinkTarget {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for TextHyperlinkTarget {
    type Error = Error;

    fn try_from(target: String) -> Result<Self> {
        Self::from_boxed(target.into_boxed_str())
    }
}

impl TryFrom<Box<str>> for TextHyperlinkTarget {
    type Error = Error;

    fn try_from(target: Box<str>) -> Result<Self> {
        Self::from_boxed(target)
    }
}

impl TryFrom<&str> for TextHyperlinkTarget {
    type Error = Error;

    fn try_from(target: &str) -> Result<Self> {
        Self::new(target)
    }
}

/// One native hyperlink attached to a nonempty UTF-16 text range.
#[allow(
    clippy::module_name_repetitions,
    reason = "TextHyperlink is the established public name of this semantic value."
)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextHyperlink {
    /// Stable semantic identity used by the owning adapter for updates.
    pub id: TextHyperlinkId,
    /// Half-open UTF-16 linked text range.
    pub range: TextRange,
    /// Native URL, mail, or application-specific target.
    pub target: TextHyperlinkTarget,
}

impl TextHyperlink {
    /// Construct a hyperlink from validated semantic components.
    #[must_use]
    pub const fn new(id: TextHyperlinkId, range: TextRange, target: TextHyperlinkTarget) -> Self {
        Self { id, range, target }
    }
}

fn validate_target(target: &str) -> Result<()> {
    if target.is_empty() {
        return Err(Error::EmptyTarget);
    }
    if target.len() > MAX_TARGET_BYTES {
        return Err(Error::TargetTooLong);
    }
    if target.trim() != target {
        return Err(Error::TargetSurroundingWhitespace);
    }
    if target.chars().any(char::is_control) {
        return Err(Error::TargetControlCharacter);
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;
    use crate::position::TextPosition;

    #[test]
    fn identifiers_are_nonzero_and_compact() {
        assert_eq!(raw::from_object_id(0), Err(Error::ZeroId));
        assert_eq!(raw::object_id(raw::from_object_id(42).unwrap()), 42);
        assert_eq!(size_of::<TextHyperlinkId>(), size_of::<u64>());
        assert_eq!(
            size_of::<Option<TextHyperlinkId>>(),
            size_of::<TextHyperlinkId>()
        );
    }

    #[test]
    fn targets_preserve_known_and_unknown_native_forms_exactly() {
        for target in [
            "https://example.com/report?q=1",
            "mailto:team@example.com",
            "?slide=next",
            "x-apple-iwork://future-target/Ä?value=%2F",
        ] {
            let target = TextHyperlinkTarget::new(target).unwrap();
            assert_eq!(target.as_str(), target.as_ref());
        }
    }

    #[test]
    fn target_validation_is_typed_and_bounded() {
        assert_eq!(TextHyperlinkTarget::new(""), Err(Error::EmptyTarget));
        assert_eq!(
            TextHyperlinkTarget::new(" https://example.com"),
            Err(Error::TargetSurroundingWhitespace)
        );
        assert_eq!(
            TextHyperlinkTarget::new("https://example.com "),
            Err(Error::TargetSurroundingWhitespace)
        );
        assert_eq!(
            TextHyperlinkTarget::new("https://example.\ncom"),
            Err(Error::TargetControlCharacter)
        );
        assert_eq!(
            TextHyperlinkTarget::new(&"x".repeat(MAX_TARGET_BYTES + 1)),
            Err(Error::TargetTooLong)
        );
        assert_eq!(
            TextHyperlinkTarget::new(&"x".repeat(MAX_TARGET_BYTES))
                .unwrap()
                .as_str()
                .len(),
            MAX_TARGET_BYTES
        );
    }

    #[test]
    fn owned_target_reuses_the_input_allocation() {
        let target = "https://example.com/one".to_owned();
        let pointer = target.as_ptr();
        let value = TextHyperlinkTarget::try_from(target).unwrap();
        assert!(std::ptr::eq(pointer, value.as_str().as_ptr()));
    }

    #[test]
    fn hyperlink_combines_a_compact_id_and_nonempty_utf16_range() {
        let id = raw::from_object_id(7).unwrap();
        let range = TextRange::new(
            TextPosition::from_utf16_code_units(2),
            TextPosition::from_utf16_code_units(7),
        )
        .unwrap();
        let hyperlink =
            TextHyperlink::new(id, range, TextHyperlinkTarget::new("?slide=next").unwrap());
        assert_eq!(hyperlink.id, id);
        assert_eq!(hyperlink.range, range);
        assert_eq!(hyperlink.target.as_str(), "?slide=next");
    }
}
