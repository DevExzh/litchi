//! Strict public types for native iWork text hyperlinks.

use crate::{Error, Result};

use litchi_iwa_text::position::TextRange;

const MAX_HYPERLINK_TARGET_BYTES: usize = 8 * 1_024;

/// Identifier of a native hyperlink smart-field object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextHyperlinkId(u64);

impl TextHyperlinkId {
    /// Construct an identifier obtained from a previously read hyperlink.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork hyperlink object identifier cannot be zero".to_owned(),
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

/// A validated native hyperlink target.
///
/// This losslessly represents ordinary URLs, `mailto:` links, and Keynote's
/// native targets such as `?slide=next` without reducing them to one scheme.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextHyperlinkTarget(Box<str>);

impl TextHyperlinkTarget {
    /// Construct a target while rejecting empty, padded, control-bearing, or
    /// unreasonably large values.
    pub fn new(target: impl Into<Box<str>>) -> Result<Self> {
        let target = target.into();
        if target.is_empty() {
            return Err(Error::ParseError(
                "iWork hyperlink target cannot be empty".to_owned(),
            ));
        }
        if target.len() > MAX_HYPERLINK_TARGET_BYTES {
            return Err(Error::ParseError(format!(
                "iWork hyperlink target exceeds {MAX_HYPERLINK_TARGET_BYTES} bytes"
            )));
        }
        if target.trim() != target.as_ref() {
            return Err(Error::ParseError(
                "iWork hyperlink target cannot have surrounding whitespace".to_owned(),
            ));
        }
        if target.chars().any(char::is_control) {
            return Err(Error::ParseError(
                "iWork hyperlink target cannot contain control characters".to_owned(),
            ));
        }
        Ok(Self(target))
    }

    /// Borrow the target exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextHyperlinkTarget {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// One native hyperlink attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextHyperlink {
    /// Stable smart-field object identifier used for update and deletion.
    pub id: TextHyperlinkId,
    /// Half-open linked text range.
    pub range: TextRange,
    /// Native URL, mail, or application-specific target.
    pub target: TextHyperlinkTarget,
}

impl TextHyperlink {
    pub(crate) fn new(id: TextHyperlinkId, range: TextRange, target: TextHyperlinkTarget) -> Self {
        Self { id, range, target }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_preserve_web_mail_and_keynote_values() {
        for target in [
            "https://example.com/report?q=1",
            "mailto:team@example.com",
            "?slide=next",
        ] {
            assert_eq!(TextHyperlinkTarget::new(target).unwrap().as_str(), target);
        }
        for invalid in ["", " https://example.com", "https://example.com\n"] {
            assert!(TextHyperlinkTarget::new(invalid).is_err());
        }
    }
}
