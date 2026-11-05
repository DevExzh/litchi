//! Typed in-memory ODT document state.

use crate::core::{Content, Meta, OwnedPackage, Styles};
use crate::elements::style::StyleRegistry;
use litchi_odf_common::{XmlStreamLimits, leniency::ToleranceReport};
use std::sync::Arc;

/// The read-only semantic projection of an `OpenDocument` text package.
///
/// A `Document` is obtained from [`crate::Package::document`]; it is owned,
/// immutable and cheap to clone (one reference-count increment). Every state
/// verb (open, edit, apply, save) lives on [`crate::Package`].
#[derive(Clone)]
pub struct Document {
    pub(super) state: Arc<State>,
}

/// The parsed parts shared by every clone of one [`Document`].
#[allow(dead_code)]
pub(crate) struct State {
    /// ZIP package containing all document files.
    pub(crate) package: OwnedPackage,
    /// Parsed `content.xml` (main document content).
    pub(crate) content: Content,
    /// Parsed `styles.xml` (document styles), if present.
    pub(crate) styles: Option<Styles>,
    /// Parsed `meta.xml` (document metadata), if present.
    pub(crate) meta: Option<Meta>,
    /// Registry of all styles in the document.
    pub(crate) style_registry: StyleRegistry,
    /// Repairs tolerated while the package was opened.
    pub(crate) tolerance: ToleranceReport,
    /// The caller-selected raw XML policy retained for reparses and edits.
    pub(crate) xml_limits: XmlStreamLimits,
    /// The structural body index, built on first use.
    pub(crate) body: std::sync::OnceLock<crate::body::Tree>,
}

impl std::fmt::Debug for Document {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Document")
            .field("package_bytes", &self.state.package.as_bytes().len())
            .finish_non_exhaustive()
    }
}
