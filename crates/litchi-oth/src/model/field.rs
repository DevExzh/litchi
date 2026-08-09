//! Inert text-field semantics.

use std::ops::Range;

/// Recognized ODF text-field family.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// `text:date`.
    Date,
    /// `text:time`.
    Time,
    /// `text:page-number`.
    PageNumber,
    /// `text:page-count`.
    PageCount,
    /// `text:title`.
    Title,
    /// `text:subject`.
    Subject,
    /// `text:author-name`.
    AuthorName,
    /// `text:user-field-get`.
    User,
    /// `text:variable-get`.
    Variable,
    /// A valid but currently unclassified text field.
    Other(String),
}

impl Kind {
    pub(crate) fn from_local(local: &[u8]) -> Self {
        match local {
            b"date" => Self::Date,
            b"time" => Self::Time,
            b"page-number" => Self::PageNumber,
            b"page-count" => Self::PageCount,
            b"title" => Self::Title,
            b"subject" => Self::Subject,
            b"author-name" => Self::AuthorName,
            b"user-field-get" => Self::User,
            b"variable-get" => Self::Variable,
            _ => Self::Other(String::from_utf8_lossy(local).into_owned()),
        }
    }
}

/// One field and its producer-stored display text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Field {
    display_range: Range<usize>,
    fixed: bool,
    kind: Kind,
    name: Option<String>,
    value: Option<String>,
}

impl Field {
    pub(crate) const fn projected(
        kind: Kind,
        name: Option<String>,
        value: Option<String>,
        fixed: bool,
        display_range: Range<usize>,
    ) -> Self {
        Self {
            display_range,
            fixed,
            kind,
            name,
            value,
        }
    }

    /// Field family. No field is evaluated by this crate.
    #[must_use]
    pub const fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Producer-visible field name, when the family carries one.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Stored value attribute, when present.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Whether the producer marked the display value fixed.
    #[must_use]
    pub const fn is_fixed(&self) -> bool {
        self.fixed
    }

    /// UTF-8 byte range of the display text in its containing block.
    #[must_use]
    pub const fn display_range(&self) -> &Range<usize> {
        &self.display_range
    }
}
