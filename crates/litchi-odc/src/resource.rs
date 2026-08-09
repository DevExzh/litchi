//! Inert package-local chart resources.

/// One non-core package member and its manifest metadata.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Resource {
    path: String,
    media_type: Option<String>,
    size: usize,
}

impl Resource {
    pub(crate) fn new(path: String, media_type: Option<String>, size: usize) -> Self {
        Self {
            path,
            media_type,
            size,
        }
    }

    #[must_use]
    pub fn path(&self) -> &str {
        &self.path
    }

    #[must_use]
    pub fn media_type(&self) -> Option<&str> {
        self.media_type.as_deref()
    }

    #[must_use]
    pub const fn size(&self) -> usize {
        self.size
    }
}
