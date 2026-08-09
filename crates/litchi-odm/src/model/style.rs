//! Inert style-catalog projection for master documents.

use std::ops::Range;

/// The package part which owns a style definition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Origin {
    /// An automatic style in `content.xml`.
    Content,
    /// A named or automatic style in `styles.xml`.
    Styles,
}

/// One named ODF style definition.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Definition {
    pub(crate) name: String,
    pub(crate) family: Option<String>,
    pub(crate) parent: Option<String>,
    pub(crate) origin: Origin,
    pub(crate) source_span: Range<usize>,
    pub(crate) name_span: Range<usize>,
}

impl Definition {
    /// Returns the native style name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Returns the native style family, when present.
    #[must_use]
    pub fn family(&self) -> Option<&str> {
        self.family.as_deref()
    }

    /// Returns the parent style name, when present.
    #[must_use]
    pub fn parent(&self) -> Option<&str> {
        self.parent.as_deref()
    }

    /// Returns the package part which owns this definition.
    #[must_use]
    pub const fn origin(&self) -> Origin {
        self.origin
    }
}
