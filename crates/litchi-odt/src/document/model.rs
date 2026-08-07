//! Typed in-memory ODT document state.

use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::elements::style::StyleRegistry;

/// An `OpenDocument` text document (`.odt`).
///
/// The document owns its validated package and parsed XML parts. The public
/// facade exposes semantic queries and atomic package edits while keeping the
/// representation private and compact.
#[allow(dead_code)]
pub struct Document {
    /// ZIP package containing all document files.
    pub(super) package: OwnedPackage,
    /// Parsed `content.xml` (main document content).
    pub(super) content: Content,
    /// Parsed `styles.xml` (document styles), if present.
    pub(super) styles: Option<Styles>,
    /// Parsed `meta.xml` (document metadata), if present.
    pub(super) meta: Option<Meta>,
    /// Registry of all styles in the document.
    pub(super) style_registry: StyleRegistry,
}
