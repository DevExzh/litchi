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

/// Bounded ordered inventory of the master body.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Structure {
    pub(crate) items: Vec<Kind>,
}

impl Structure {
    /// Returns direct body children in authored document order.
    #[must_use]
    pub fn items(&self) -> &[Kind] {
        &self.items
    }
}
