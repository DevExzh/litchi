//! Legacy DOC header/footer model codec.
//!
//! The document parser already decodes the story character range and
//! paragraphs before constructing [`HeaderFooter`]. This boundary therefore
//! performs the lossless package-payload-to-model transfer without copying
//! either owned collection.

use super::model::HeaderFooter;
use super::package::Story;

/// Materialize a typed semantic story from package-layer data.
pub(super) fn decode(story: Story) -> HeaderFooter {
    HeaderFooter {
        header_footer_type: story.header_footer_type,
        text: story.text,
        paragraphs: story.paragraphs,
    }
}
