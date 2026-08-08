//! Bounded inert syntax retained when the semantic codec does not understand it.

/// Semantic owner of unsupported syntax that cannot be safely reconstructed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Context {
    /// Font, color, style, document-information, or other header metadata.
    Metadata,
    /// A header or footer story.
    HeaderFooter,
    /// A field instruction or result story. Fields remain inert.
    Field,
    /// A footnote or endnote story.
    Note,
    /// A table row, cell, or nested-table property group.
    Table,
    /// Picture, object-result, or drawing-owned syntax.
    Drawing,
    /// Revision-owned syntax.
    Review,
    /// Another structural owner not modeled by the canonical writer.
    Other,
}

/// Where an opaque syntax node belongs in the retained document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Anchor {
    /// Syntax anchored at a UTF-8 byte boundary in the visible body story.
    Body(usize),
    /// Exact lexical position of syntax whose structural owner is not writable.
    Structural {
        /// Typed semantic context active at the source location.
        context: Context,
        /// Zero-based token position in the validated lexical stream.
        token: usize,
        /// Group depth at which the syntax was encountered.
        depth: usize,
    },
}

/// The unsupported RTF construct retained by a [`Node`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// A complete unsupported destination group, including its braces.
    Destination,
    /// One otherwise-unhandled control word, including its delimiter.
    ControlWord,
}

/// A validated, inert fragment of unsupported RTF syntax.
///
/// Nodes are produced only by the parser. Their bytes are never interpreted as
/// fields, objects, macros, controls, or executable content.
#[derive(Debug, PartialEq, Eq)]
pub struct Node {
    kind: Kind,
    anchor: Anchor,
    source: Vec<u8>,
}

impl Node {
    pub(crate) fn new(kind: Kind, anchor: Anchor, source: Vec<u8>) -> Self {
        Self {
            kind,
            anchor,
            source,
        }
    }

    /// Return the unsupported construct category.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the semantic location at which the syntax was encountered.
    #[must_use]
    pub const fn anchor(&self) -> Anchor {
        self.anchor
    }

    /// Borrow the exact uncompressed RTF transport bytes.
    #[must_use]
    pub fn source(&self) -> &[u8] {
        &self.source
    }
}
