//! The glossary batch text family of a source-backed package (record
//! 0355a): replace the text of direct paragraphs of several entries of the
//! glossary document in one atomic change.
//!
//! [`Package::edit_glossary_text_batch`](crate::source_backed::Package::edit_glossary_text_batch)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the selected story part.

use litchi_core::Position;

use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::story_text::{self as record, GlossaryBatchSelector};

/// The selected-story record patch the family publishes.
#[derive(Clone)]
pub(crate) struct GlossaryTextChange(record::GlossaryBatchPatch);

impl Change for GlossaryTextChange {
    type Error = Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package, Error> {
        let (part, bytes) = package.glossary_text_batch_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        let base = package.clone();
        Ok(package.deferred(move |writer, policy| {
            base.write_part_overlay(writer, &part, &bytes, policy)
        })?)
    }
}

source_family!("glossary batch text", GlossaryTextChange, record::Error);

impl From<record::GlossaryBatchPatch> for Patch {
    /// Replay a selected-story record patch onto a package whose selected
    /// story holds its before bytes.
    fn from(patch: record::GlossaryBatchPatch) -> Self {
        Self {
            change: GlossaryTextChange(patch),
            restore: None,
        }
    }
}

/// The budgets of this family, reported by [`Error::Limit`].
pub use crate::source_backed::story_text::LimitKind;

/// A staged text change of one selected story. The source is never changed:
/// [`Self::commit`] returns a new package.
pub struct Edit {
    package: Package,
    inner: record::GlossaryBatchEdit,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The selected story with the staged changes.
    ///
    /// # Errors
    ///
    /// Returns a projection error.
    pub fn projected(&self) -> Result<record::GlossaryBatchSnapshot, Error> {
        self.inner.projected()
    }

    /// Replace all text in one direct paragraph of one selected entry.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal or limit error, without staging anything.
    pub fn replace_paragraph_text(
        &mut self,
        entry_index: usize,
        position: Position,
        authored_text: impl AsRef<str>,
    ) -> Result<&mut Self, Error> {
        self.inner
            .replace_paragraph_text(entry_index, position, authored_text)?;
        Ok(self)
    }

    /// Commit the staged text as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source, topology, output-limit or signature error
    /// before any output.
    pub fn commit(self) -> Result<Commit, Error> {
        let patch = self.inner.finish()?;
        commit_change(&self.package, GlossaryTextChange(patch))
    }
}

impl Package {
    /// Stage a text change of one selected story.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::glossary_text_batch_snapshot`].
    pub fn edit_glossary_text_batch(&self, selector: GlossaryBatchSelector) -> Result<Edit, Error> {
        Ok(Edit {
            package: self.clone(),
            inner: self.glossary_text_batch_snapshot(selector)?.edit_engine()?,
        })
    }

    /// Replay a selected-story patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] or [`Error::TopologyConflict`] for
    /// another story state.
    pub fn apply_glossary_text_batch(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_change(self, patch)
    }
}
