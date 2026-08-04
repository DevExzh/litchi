//! Semantic values for PresentationML legacy comments.
//!
//! These names are scoped by the \`comments\` module; compatibility aliases
//! retain the historical public vocabulary at the module boundary.

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

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub author_id: u32,
    pub date_time: Option<String>,
    pub index: u32,
    pub x: i64,
    pub y: i64,
    pub text: String,
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
