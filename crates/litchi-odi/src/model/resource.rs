//! Package-local image resource inventory.

/// A package-local image reference discovered in `content.xml`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    frame: usize,
    href: String,
    path: String,
    media_type: Option<String>,
    present: bool,
}

impl Resource {
    pub(crate) fn new(
        frame: usize,
        href: String,
        path: String,
        media_type: Option<String>,
        present: bool,
    ) -> Self {
        Self {
            frame,
            href,
            path,
            media_type,
            present,
        }
    }

    /// Returns the zero-based semantic frame position that references this resource.
    #[must_use]
    pub fn frame(&self) -> usize {
        self.frame
    }

    /// Returns the original inert `xlink:href` value.
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Returns the safely resolved package path.
    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    /// Returns the manifest media type, if declared.
    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    /// Returns whether the referenced package member exists.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.present
    }
}
