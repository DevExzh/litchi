//! Package-neutral values for the Word 2012 `footnoteColumns` extension.

use crate::error::{Error, Result};
use std::fmt;

/// Footnote-area layout requested for a section.
///
/// A value of zero is meaningful: `[MS-DOCX]` specifies that zero restores
/// the page's ordinary column-derived footnote layout. Absence is retained
/// separately by [`super::Snapshot`], so callers can distinguish an omitted
/// extension from an explicit zero.
#[derive(Debug, Clone, Copy, Default, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct Layout {
    columns: i32,
}

impl Layout {
    /// Construct a layout from the decimal column count.
    pub fn new(columns: i32) -> Result<Self> {
        if columns < 0 {
            return Err(Error::InvalidFormat(
                "footnote column count cannot be negative".into(),
            ));
        }
        Ok(Self { columns })
    }

    /// Return the requested number of footnote columns.
    #[must_use]
    pub const fn columns(self) -> i32 {
        self.columns
    }

    /// Whether Word should derive the footnote layout from the displayed page.
    #[must_use]
    pub const fn follows_page_layout(self) -> bool {
        self.columns == 0
    }
}

impl TryFrom<i32> for Layout {
    type Error = Error;

    #[inline]
    fn try_from(columns: i32) -> Result<Self> {
        Self::new(columns)
    }
}

impl fmt::Display for Layout {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.columns.fmt(formatter)
    }
}
