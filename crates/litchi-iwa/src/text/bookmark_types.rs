//! Strict public value types for native ranged iWork bookmarks.

use crate::{Error, Result};

use litchi_iwa_text::position::TextRange;

const MAX_BOOKMARK_NAME_BYTES: usize = 1_024;

/// Identifier of a native bookmark-field object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextBookmarkId(u64);

impl TextBookmarkId {
    /// Construct an identifier obtained from a previously read bookmark.
    pub fn from_object_id(identifier: u64) -> Result<Self> {
        if identifier == 0 {
            return Err(Error::ParseError(
                "iWork bookmark object identifier cannot be zero".to_owned(),
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

/// Optional display name stored on a native bookmark.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct TextBookmarkName(Box<str>);

impl TextBookmarkName {
    /// Validate a nonempty, unpadded bookmark name.
    pub fn new(name: impl Into<Box<str>>) -> Result<Self> {
        let name = name.into();
        if name.is_empty() {
            return Err(Error::ParseError(
                "iWork bookmark name cannot be empty".to_owned(),
            ));
        }
        if name.len() > MAX_BOOKMARK_NAME_BYTES {
            return Err(Error::ParseError(format!(
                "iWork bookmark name exceeds {MAX_BOOKMARK_NAME_BYTES} bytes"
            )));
        }
        if name.trim() != name.as_ref() {
            return Err(Error::ParseError(
                "iWork bookmark name cannot have surrounding whitespace".to_owned(),
            ));
        }
        if name.chars().any(char::is_control) {
            return Err(Error::ParseError(
                "iWork bookmark name cannot contain control characters".to_owned(),
            ));
        }
        Ok(Self(name))
    }

    /// Borrow the name exactly as stored by iWork.
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for TextBookmarkName {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Visibility flag stored on a bookmark field.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum TextBookmarkVisibility {
    #[default]
    Visible,
    Hidden,
    /// A value written by a newer iWork version.
    Unknown(u32),
}

impl TextBookmarkVisibility {
    pub(crate) const fn from_raw(raw: u32) -> Self {
        match raw {
            0 => Self::Visible,
            1 => Self::Hidden,
            unknown => Self::Unknown(unknown),
        }
    }

    pub(crate) const fn as_raw(self) -> u32 {
        match self {
            Self::Visible => 0,
            Self::Hidden => 1,
            Self::Unknown(raw) => raw,
        }
    }
}

/// Writable bookmark object settings.
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct TextBookmarkSettings {
    pub name: Option<TextBookmarkName>,
    pub visibility: TextBookmarkVisibility,
}

impl TextBookmarkSettings {
    /// Create visible unnamed settings matching Pages' default bookmark action.
    pub const fn new() -> Self {
        Self {
            name: None,
            visibility: TextBookmarkVisibility::Visible,
        }
    }

    /// Set a validated display name.
    pub fn with_name(mut self, name: TextBookmarkName) -> Self {
        self.name = Some(name);
        self
    }

    /// Set the native visibility flag.
    pub const fn with_visibility(mut self, visibility: TextBookmarkVisibility) -> Self {
        self.visibility = visibility;
        self
    }
}

/// One native bookmark attached to a nonempty UTF-16 text range.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct TextBookmark {
    pub id: TextBookmarkId,
    pub range: TextRange,
    pub settings: TextBookmarkSettings,
}

impl TextBookmark {
    pub(crate) fn new(
        id: TextBookmarkId,
        range: TextRange,
        settings: TextBookmarkSettings,
    ) -> Self {
        Self {
            id,
            range,
            settings,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_and_identifiers_are_strict() {
        assert_eq!(
            TextBookmarkName::new("Methods").unwrap().as_str(),
            "Methods"
        );
        for invalid in ["", " padded", "padded ", "line\nbreak"] {
            assert!(TextBookmarkName::new(invalid).is_err());
        }
        assert!(TextBookmarkId::from_object_id(0).is_err());
        assert_eq!(TextBookmarkId::from_object_id(7).unwrap().object_id(), 7);
    }

    #[test]
    fn visibility_preserves_future_values() {
        for raw in [0, 1, 7] {
            assert_eq!(TextBookmarkVisibility::from_raw(raw).as_raw(), raw);
        }
    }
}
