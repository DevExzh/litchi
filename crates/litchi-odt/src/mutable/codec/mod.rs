//! Contextual facade for mutable ODT snapshot edits.
//!
//! The public `MutableDocument` methods live in [`semantic`]. XML snapshot
//! access and package-state transitions stay behind the small seams in
//! [`xml`] and [`package`].

mod package;
mod semantic;
mod xml;

use super::model::MutableDocument;
use litchi_core::Result;

impl MutableDocument {
    /// Run a read-only operation against the authoritative content snapshot.
    ///
    /// This facade remains `pub(super)` because sibling mutable owners use the
    /// same snapshot semantics without depending on a nested codec module.
    pub(super) fn with_content_xml<T>(
        &self,
        operation: impl FnOnce(&str) -> Result<T>,
    ) -> Result<T> {
        package::with_content_xml(self, operation)
    }

    /// Apply one lossless transformation to the authoritative content
    /// snapshot and commit it only after the transformation succeeds.
    pub(super) fn edit_content_xml(
        &mut self,
        operation: impl FnOnce(&str) -> Result<String>,
    ) -> Result<()> {
        let updated = self.with_content_xml(operation)?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Discard the inline XML snapshot after a structural model mutation.
    pub(super) fn invalidate_content_xml(&mut self) {
        package::invalidate_content_xml(self);
    }
}
