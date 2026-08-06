//! Lossless bytes for metadata streams outside the typed DOC inventory.

use std::sync::Arc;

/// An inert, lossless stream that was not decoded as a known DOC/OLE stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Unknown {
    path: Vec<String>,
    bytes: Arc<[u8]>,
}

impl Unknown {
    /// The exact path relative to the selected ObjectPool storage.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// The final CFB stream name, when the path is non-empty.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.path.last().map(String::as_str)
    }

    /// The exact stream bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Shared ownership of the exact stream allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    pub(in crate::embedded_object) fn from_parts(path: &[String], bytes: Arc<[u8]>) -> Self {
        Self {
            path: path.to_vec(),
            bytes,
        }
    }
}
