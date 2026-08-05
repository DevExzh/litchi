//! Package-layer payload for a legacy DOC header/footer story.
//!
//! The enclosing document owns FIB/header-table traversal. This narrow value
//! is the owner boundary used to transfer already-decoded story data into the
//! semantic model without exposing stream or table details in its API.

use crate::paragraph::Paragraph;
use crate::parts::headers::HeaderFooterType;

/// Decoded story payload supplied by the legacy DOC package reader.
pub(super) struct Story {
    pub(super) header_footer_type: HeaderFooterType,
    pub(super) text: String,
    pub(super) paragraphs: Vec<Paragraph>,
}

impl Story {
    pub(super) fn new(
        header_footer_type: HeaderFooterType,
        text: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            header_footer_type,
            text,
            paragraphs,
        }
    }
}
