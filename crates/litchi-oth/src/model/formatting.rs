//! Inline formatting semantics.

use std::ops::Range;

/// A styled character range within one projected text block.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Run {
    range: Range<usize>,
    style_name: String,
}

impl Run {
    pub(crate) const fn projected(range: Range<usize>, style_name: String) -> Self {
        Self { range, style_name }
    }

    /// UTF-8 byte range in the block's projected text.
    #[must_use]
    pub const fn range(&self) -> &Range<usize> {
        &self.range
    }

    /// Referenced ODF text style.
    #[must_use]
    pub fn style_name(&self) -> &str {
        &self.style_name
    }
}
