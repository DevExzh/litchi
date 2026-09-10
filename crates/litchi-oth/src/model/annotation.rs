//! Read-only projection of ODF office annotations.

/// One inert `office:annotation` attached to the text body.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Annotation {
    creator: Option<String>,
    date: Option<String>,
    date_string: Option<String>,
    display: Option<bool>,
    initials: Option<String>,
    name: Option<String>,
    text: String,
}

impl Annotation {
    pub(crate) const fn projected(
        name: Option<String>,
        creator: Option<String>,
        date: Option<String>,
        date_string: Option<String>,
        initials: Option<String>,
        display: Option<bool>,
        text: String,
    ) -> Self {
        Self {
            creator,
            date,
            date_string,
            display,
            initials,
            name,
            text,
        }
    }

    /// Optional office annotation name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Annotation creator, if present.
    #[must_use]
    pub fn creator(&self) -> Option<&str> {
        self.creator.as_deref()
    }

    /// Machine-readable annotation date, if present.
    #[must_use]
    pub fn date(&self) -> Option<&str> {
        self.date.as_deref()
    }

    /// Human-readable annotation date string, if present.
    #[must_use]
    pub fn date_string(&self) -> Option<&str> {
        self.date_string.as_deref()
    }

    /// Producer initials, if present.
    #[must_use]
    pub fn initials(&self) -> Option<&str> {
        self.initials.as_deref()
    }

    /// Requested display state, if present.
    #[must_use]
    pub const fn display(&self) -> Option<bool> {
        self.display
    }

    /// Inert visible annotation text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}
