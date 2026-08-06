//! Mutable document state and the in-memory structural projection.

use crate::core::OwnedPackage;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use litchi_core::Metadata;

/// Document element type used to retain top-level insertion order.
#[derive(Debug, Clone)]
pub(super) enum DocumentElement {
    /// A paragraph element.
    Paragraph(Paragraph),
    /// A heading element.
    Heading(Heading),
    /// A table element.
    Table(Table),
    /// A text list element.
    List(List),
    /// A standalone drawing frame at body level.
    Frame(crate::elements::element::Element),
}

/// A mutable ODT document that supports structural and lossless XML edits.
pub struct MutableDocument {
    /// Document elements in insertion order.
    pub(super) elements: Vec<DocumentElement>,
    /// Document metadata.
    pub(super) metadata: Metadata,
    /// Root MIME type written on save.
    pub(super) mimetype: String,
    /// Retained `styles.xml`, when present.
    pub(super) styles_xml: Option<String>,
    /// Original package used to retain auxiliary parts during rewriting.
    pub(super) source_package: Option<OwnedPackage>,
    /// Authoritative content XML used by byte-preserving inline mutations.
    pub(super) content_xml: Option<String>,
    /// Authored picture payloads written into the package on save.
    pub(super) pending_images: Vec<crate::frame::Part>,
    /// Monotonic counter for authored frame names.
    pub(super) next_frame_number: usize,
}

impl MutableDocument {
    /// Create a new empty mutable document.
    ///
    /// ```
    /// use litchi_odt::mutable::MutableDocument;
    ///
    /// let document = MutableDocument::new();
    /// assert!(document.paragraphs().is_empty());
    /// ```
    pub fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.text".to_string(),
            styles_xml: None,
            source_package: None,
            content_xml: None,
            pending_images: Vec::new(),
            next_frame_number: 1,
        }
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}
