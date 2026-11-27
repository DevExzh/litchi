//! Inert package-local drawing resource inventory.

/// A package-local image or media member referenced by drawing XML.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    occurrence: usize,
    href: String,
    path: String,
    media_type: Option<String>,
    present: bool,
}

impl Resource {
    pub(crate) fn new(
        occurrence: usize,
        href: String,
        path: String,
        media_type: Option<String>,
        present: bool,
    ) -> Self {
        Self {
            occurrence,
            href,
            path,
            media_type,
            present,
        }
    }

    /// Zero-based image occurrence in the bounded drawing media scan.
    #[must_use]
    pub const fn occurrence(&self) -> usize {
        self.occurrence
    }

    /// Original inert `xlink:href` value.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Safely resolved package member path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Manifest media type, when declared.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Whether the referenced member exists in the package.
    #[must_use]
    pub const fn is_present(&self) -> bool {
        self.present
    }
}
