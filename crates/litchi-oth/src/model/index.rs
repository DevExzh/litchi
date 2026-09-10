//! Read-only projection of generated ODF text indexes.

/// Generated text-index family.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    TableOfContents,
    Illustration,
    Table,
    Object,
    User,
    Alphabetical,
    Bibliography,
    Other(String),
}

impl Kind {
    /// Lexical ODF index element name represented by this kind.
    #[must_use]
    pub fn as_str(&self) -> &str {
        match self {
            Self::TableOfContents => "table-of-content",
            Self::Illustration => "illustration-index",
            Self::Table => "table-index",
            Self::Object => "object-index",
            Self::User => "user-index",
            Self::Alphabetical => "alphabetical-index",
            Self::Bibliography => "bibliography",
            Self::Other(value) => value,
        }
    }
}

/// One generated index declaration with its inert cached body text.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Index {
    body: String,
    kind: Kind,
    name: Option<String>,
    protected: bool,
    source: Option<String>,
}

impl Index {
    pub(crate) const fn projected(
        kind: Kind,
        name: Option<String>,
        protected: bool,
        source: Option<String>,
        body: String,
    ) -> Self {
        Self {
            body,
            kind,
            name,
            protected,
            source,
        }
    }

    /// Index family.
    #[must_use]
    pub const fn kind(&self) -> &Kind {
        &self.kind
    }

    /// Optional producer-visible index name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Whether the index is marked protected.
    #[must_use]
    pub const fn is_protected(&self) -> bool {
        self.protected
    }

    /// Inert source/template text, if a source element was present.
    #[must_use]
    pub fn source(&self) -> Option<&str> {
        self.source.as_deref()
    }

    /// Stored cached index body text. Entries are never regenerated.
    #[must_use]
    pub fn body(&self) -> &str {
        &self.body
    }
}
