//! Ordered master-body structure semantics.

use litchi_core::Position;

/// A generated master-document index kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum IndexKind {
    TableOfContents,
    Illustration,
    Table,
    Object,
    User,
    Alphabetical,
    Bibliography,
}

/// One direct child of the `office:text` master body.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    Paragraph,
    Heading,
    List,
    Table,
    Section(Position),
    GeneratedIndex(IndexKind),
    Declarations,
    Other,
}

/// One generated index declared directly in the master body.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct GeneratedIndex {
    pub(crate) item: Position,
    pub(crate) kind: IndexKind,
    pub(crate) name: Option<String>,
    pub(crate) xml_id: Option<String>,
}

impl GeneratedIndex {
    /// Returns the position in [`Structure::items`].
    #[must_use]
    pub const fn item(&self) -> Position {
        self.item
    }

    /// Returns the generated index kind.
    #[must_use]
    pub const fn kind(&self) -> IndexKind {
        self.kind
    }

    /// Returns the optional `text:name` identity.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns the optional `xml:id` identity.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }
}

/// Bounded ordered inventory of the master body.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Structure {
    pub(crate) items: Vec<Kind>,
    pub(crate) generated_indexes: Vec<GeneratedIndex>,
}

impl Structure {
    /// Returns direct body children in authored document order.
    #[must_use]
    pub fn items(&self) -> &[Kind] {
        &self.items
    }

    /// Returns generated indexes with their body position and identities.
    #[must_use]
    pub fn generated_indexes(&self) -> &[GeneratedIndex] {
        &self.generated_indexes
    }
}
