//! Mutable document state and the in-memory structural projection.

use crate::core::OwnedPackage;
use crate::elements::table::Table;
use crate::elements::text::{Heading, List, Paragraph};
use litchi_core::Metadata;
use litchi_odf_common::XmlStreamLimits;

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
pub(crate) struct MutableDocument {
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
    /// Retained raw XML policy from the immutable owner.
    pub(super) xml_limits: XmlStreamLimits,
    /// Monotonic counter for authored frame names.
    pub(super) next_frame_number: usize,
}

impl MutableDocument {
    /// Create a new empty mutable document.
    pub(crate) fn new() -> Self {
        Self {
            elements: Vec::new(),
            metadata: Metadata::default(),
            mimetype: "application/vnd.oasis.opendocument.text".to_string(),
            styles_xml: None,
            source_package: None,
            content_xml: None,
            pending_images: Vec::new(),
            xml_limits: XmlStreamLimits::default(),
            next_frame_number: 1,
        }
    }
}

impl Default for MutableDocument {
    fn default() -> Self {
        Self::new()
    }
}
