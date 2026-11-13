//! Mutable document metadata accessors.

use super::model::MutableDocument;
#[cfg(test)]
use litchi_core::Metadata;

impl MutableDocument {
    /// Get a mutable reference to the document metadata.
    #[cfg(test)]
    pub(crate) fn metadata_mut(&mut self) -> &mut Metadata {
        &mut self.metadata
    }
}
