//! Master-document section semantics.

use super::subdocument::Subdocument;
use litchi_core::Position;
use std::ops::Range;

/// A master-document section.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Section {
    name: String,
    children: Vec<Subdocument>,
}

/// One parsed `text:section` in document order.
///
/// Parent and child positions index the immutable [`Tree::sections`] slice.
/// A linked section additionally carries the position of its entry in
/// [`crate::Master::subdocuments`].
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Node {
    pub(crate) name: String,
    pub(crate) style_name: Option<String>,
    pub(crate) xml_id: Option<String>,
    pub(crate) protected: Option<bool>,
    pub(crate) parent: Option<Position>,
    pub(crate) children: Vec<Position>,
    pub(crate) reference: Option<Position>,
    pub(crate) local_reference: Option<Position>,
    pub(crate) dde_source: bool,
    pub(crate) source_span: Range<usize>,
    pub(crate) name_span: Range<usize>,
}

impl Node {
    /// Returns the unique section name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the section style name, when present.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Returns the native XML identity, when present.
    #[must_use]
    pub fn xml_id(&self) -> Option<&str> {
        self.xml_id.as_deref()
    }

    /// Returns the parsed `text:protected` value, when explicitly present.
    #[must_use]
    pub const fn protected(&self) -> Option<bool> {
        self.protected
    }

    /// Returns the parent section position, or `None` for a root section.
    #[must_use]
    pub const fn parent(&self) -> Option<Position> {
        self.parent
    }

    /// Returns direct child section positions in document order.
    #[must_use]
    pub fn children(&self) -> &[Position] {
        &self.children
    }

    /// Returns the linked-subdocument position, when this is a linked section.
    #[must_use]
    pub const fn reference(&self) -> Option<Position> {
        self.reference
    }

    /// Returns the position in [`Tree::local_references`] when the section
    /// links to another section in this master document.
    #[must_use]
    pub const fn local_reference(&self) -> Option<Position> {
        self.local_reference
    }

    /// Reports whether this section declares an inert `office:dde-source`.
    #[must_use]
    pub const fn has_dde_source(&self) -> bool {
        self.dde_source
    }
}

/// One `text:section-source text:section-name` relationship.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LocalReference {
    pub(crate) owner: Position,
    pub(crate) target_name: String,
    pub(crate) target: Option<Position>,
}

impl LocalReference {
    /// Returns the section which owns the source declaration.
    #[must_use]
    pub const fn owner(&self) -> Position {
        self.owner
    }

    /// Returns the referenced section name exactly as authored.
    #[must_use]
    pub fn target_name(&self) -> &str {
        &self.target_name
    }

    /// Resolves the target when that section exists in this master.
    #[must_use]
    pub const fn target(&self) -> Option<Position> {
        self.target
    }
}

/// Bounded immutable section tree projected from `content.xml`.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Tree {
    pub(crate) sections: Vec<Node>,
    pub(crate) roots: Vec<Position>,
    pub(crate) local_references: Vec<LocalReference>,
}

impl Tree {
    /// Returns every section in document order.
    #[must_use]
    pub fn sections(&self) -> &[Node] {
        &self.sections
    }

    /// Returns root section positions in document order.
    #[must_use]
    pub fn roots(&self) -> &[Position] {
        &self.roots
    }

    /// Returns local section relationships in document order.
    #[must_use]
    pub fn local_references(&self) -> &[LocalReference] {
        &self.local_references
    }

    /// Resolves a checked section position.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<&Node> {
        self.sections.get(position.get())
    }
}

impl Section {
    pub fn new(name: impl Into<String>) -> Self {
        Self {
            name: name.into(),
            children: Vec::new(),
        }
    }

    /// Returns the section name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the subdocuments contained in the section.
    #[must_use]
    pub fn children(&self) -> &[Subdocument] {
        &self.children
    }

    pub fn push(&mut self, child: Subdocument) {
        self.children.push(child);
    }
}
