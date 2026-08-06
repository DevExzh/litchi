//! Mutable document metadata accessors.

use super::model::MutableDocument;
use litchi_core::Metadata;

impl MutableDocument {
    /// Get the document metadata.
    pub fn metadata(&self) -> &Metadata {
        &self.metadata
    }

    /// Get a mutable reference to the document metadata.
    pub fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }
}
