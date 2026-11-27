//! Semantic values for `PresentationML` legacy comments.

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => super::PML,
            Self::Strict => super::STRICT_PML,
        }
    }

    pub(crate) fn comments_relationship(self) -> &'static str {
        match self {
            Self::Transitional => super::COMMENTS_REL,
            Self::Strict => super::STRICT_COMMENTS_REL,
        }
    }

    pub(crate) fn authors_relationship(self) -> &'static str {
        match self {
            Self::Transitional => super::AUTHORS_REL,
            Self::Strict => super::STRICT_AUTHORS_REL,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Author {
    pub id: u32,
    pub name: String,
    pub initials: String,
    pub last_index: u32,
    pub color_index: u32,
}

impl Author {
    /// Create an author with the standard initial comment metadata.
    pub fn new(id: u32, name: impl Into<String>, initials: impl Into<String>) -> Self {
        Self {
            id,
            name: name.into(),
            initials: initials.into(),
            last_index: 0,
            color_index: id % 6,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub author_id: u32,
    pub date_time: Option<String>,
    pub index: u32,
    pub x: i64,
    pub y: i64,
    pub text: String,
}

impl Comment {
    /// Create a comment with the first valid per-author index.
    pub fn new(author_id: u32, text: impl Into<String>, x: i64, y: i64) -> Self {
        Self {
            author_id,
            date_time: None,
            index: 1,
            x,
            y,
            text: text.into(),
        }
    }

    #[must_use]
    pub fn with_date_time(mut self, value: impl Into<String>) -> Self {
        self.date_time = Some(value.into());
        self
    }

    #[must_use]
    pub fn with_index(mut self, value: u32) -> Self {
        self.index = value;
        self
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct List {
    pub slide_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub comments: Vec<Comment>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comments {
    pub author_relationship_id: String,
    pub author_part_name: String,
    pub authors: Vec<Author>,
    pub slides: Vec<List>,
}
