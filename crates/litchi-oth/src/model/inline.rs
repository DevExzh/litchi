//! Detached rich inline content for text blocks.

/// One safely authorable inline item.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Content {
    /// A point bookmark at the current text offset.
    BookmarkPoint(String),
    /// The end marker of a named bookmark range.
    BookmarkRangeEnd(String),
    /// The start marker of a named bookmark range.
    BookmarkRangeStart(String),
    /// An inert field with stored display text.
    Field(Field),
    /// An inert hyperlink.
    Link(crate::link::Link),
    /// A styled text span.
    Span(Span),
    /// Plain character data.
    Text(String),
}

impl Content {
    /// Creates plain inline text.
    #[must_use]
    pub fn text(value: impl Into<String>) -> Self {
        Self::Text(value.into())
    }

    /// Creates a point bookmark.
    #[must_use]
    pub fn bookmark(name: impl Into<String>) -> Self {
        Self::BookmarkPoint(name.into())
    }

    /// Creates a bookmark range start marker.
    #[must_use]
    pub fn bookmark_start(name: impl Into<String>) -> Self {
        Self::BookmarkRangeStart(name.into())
    }

    /// Creates a bookmark range end marker.
    #[must_use]
    pub fn bookmark_end(name: impl Into<String>) -> Self {
        Self::BookmarkRangeEnd(name.into())
    }
}

/// A detached inert field value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Field {
    display: String,
    fixed: bool,
    kind: crate::field::Kind,
    name: Option<String>,
    value: Option<String>,
}

impl Field {
    /// Creates a field with stored display text. It is never evaluated.
    #[must_use]
    pub fn new(kind: crate::field::Kind, display: impl Into<String>) -> Self {
        Self {
            display: display.into(),
            fixed: false,
            kind,
            name: None,
            value: None,
        }
    }

    /// Sets the producer-visible field name.
    #[must_use]
    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    /// Sets the inert producer-stored value attribute.
    #[must_use]
    pub fn with_value(mut self, value: impl Into<String>) -> Self {
        self.value = Some(value.into());
        self
    }

    /// Marks the stored display value fixed.
    #[must_use]
    pub const fn fixed(mut self) -> Self {
        self.fixed = true;
        self
    }

    /// Stored display text.
    #[must_use]
    pub fn display(&self) -> &str {
        &self.display
    }

    /// Field family.
    #[must_use]
    pub const fn kind(&self) -> &crate::field::Kind {
        &self.kind
    }

    /// Producer-visible name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Producer-stored value.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Whether the display is fixed.
    #[must_use]
    pub const fn is_fixed(&self) -> bool {
        self.fixed
    }
}

/// A detached styled text span.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Span {
    style_name: String,
    text: String,
}

impl Span {
    /// Creates a styled span.
    #[must_use]
    pub fn new(style_name: impl Into<String>, text: impl Into<String>) -> Self {
        Self {
            style_name: style_name.into(),
            text: text.into(),
        }
    }

    /// Referenced text style.
    #[must_use]
    pub fn style_name(&self) -> &str {
        &self.style_name
    }

    /// Span text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
