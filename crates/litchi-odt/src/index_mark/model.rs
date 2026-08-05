//! Semantic index source-mark values and fragment types.

use crate::index::TextIndexAttribute;

/// An index source-mark family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextIndexMarkKind {
    TableOfContents,
    User,
    Alphabetical,
    Bibliography,
}

/// A point or resolved range mark that contributes an entry to a generated index.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextIndexMark {
    pub(super) kind: TextIndexMarkKind,
    pub(super) id: Option<String>,
    pub(super) value: String,
    pub(super) range: bool,
    pub(super) attributes: Vec<TextIndexAttribute>,
}

impl TextIndexMark {
    pub fn kind(&self) -> TextIndexMarkKind {
        self.kind
    }

    pub fn id(&self) -> Option<&str> {
        self.id.as_deref()
    }

    /// Point marks return their stored string; range marks return their referenced visible text.
    pub fn value(&self) -> &str {
        &self.value
    }

    pub fn is_range(&self) -> bool {
        self.range
    }

    pub fn attributes(&self) -> &[TextIndexAttribute] {
        &self.attributes
    }

    pub fn attribute(&self, namespace_uri: Option<&str>, local_name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|attribute| {
                attribute.namespace_uri() == namespace_uri && attribute.local_name() == local_name
            })
            .map(TextIndexAttribute::value)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct TextAlphabeticalMarkMetadata {
    pub key1: Option<String>,
    pub key2: Option<String>,
    pub string_value_phonetic: Option<String>,
    pub key1_phonetic: Option<String>,
    pub key2_phonetic: Option<String>,
    pub main_entry: Option<bool>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TextIndexMarkFragments {
    Point(String),
    Range { start: String, end: String },
}
