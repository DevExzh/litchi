//! The selected-story text family of a source-backed package (record
//! 0355a): replace the text of direct paragraphs in the main story or one
//! header, footer, footnote, endnote, comment or glossary entry.
//!
//! [`Package::edit_story_text`](crate::source_backed::Package::edit_story_text)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the selected story part.

use std::cell::{RefCell, RefMut};

use litchi_core::Position;

use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::story_text::{self as record, Selector};

/// The selected-story record patch the family publishes.
#[derive(Clone)]
pub(crate) struct StoryTextChange(record::Patch);

impl Change for StoryTextChange {
    type Error = Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package, Error> {
        let (part, bytes) = package.story_text_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        Ok(package.deferred_part_overlay(part, bytes)?)
    }
}

source_family!("selected-story text", StoryTextChange, record::Error);

impl From<record::Patch> for Patch {
    /// Replay a selected-story record patch onto a package whose selected
    /// story holds its before bytes.
    fn from(patch: record::Patch) -> Self {
        Self {
            change: StoryTextChange(patch),
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
    selector: Selector,
    limits: record::Limits,
    inner: RefCell<Option<record::Edit>>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn materialize(&self) -> Result<record::Edit, Error> {
        self.package
            .story_text_snapshot_bounded(self.selector.clone(), self.limits)?
            .edit_engine()
    }

    fn ensure_inner(&self) -> Result<(), Error> {
        let mut inner = self.inner.borrow_mut();
        if inner.is_none() {
            *inner = Some(self.materialize()?);
        }
        Ok(())
    }

    fn inner_mut(&self) -> Result<RefMut<'_, record::Edit>, Error> {
        self.ensure_inner()?;
        RefMut::filter_map(self.inner.borrow_mut(), Option::as_mut).map_err(|_| {
            Error::Document(crate::Error::internal(
                "story_text.edit",
                "story-text edit materializer is unavailable",
            ))
        })
    }

    fn into_inner(mut self) -> Result<record::Edit, Error> {
        if self.inner.get_mut().is_none() {
            let inner = self.materialize()?;
            *self.inner.get_mut() = Some(inner);
        }
        self.inner.into_inner().ok_or_else(|| {
            Error::Document(crate::Error::internal(
                "story_text.edit",
                "story-text edit materializer is unavailable",
            ))
        })
    }

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
    pub fn projected(&self) -> Result<record::Snapshot, Error> {
        self.inner_mut()?.projected()
    }

    /// Replace all text in one direct paragraph of the selected story.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal or limit error, without staging anything.
    pub fn replace_paragraph_text(
        &mut self,
        position: Position,
        authored_text: impl AsRef<str>,
    ) -> Result<&mut Self, Error> {
        self.inner_mut()?
            .replace_paragraph_text(position, authored_text)?;
        Ok(self)
    }

    /// Commit the staged text as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source, topology, output-limit or signature error
    /// before any output.
    pub fn commit(self) -> Result<Commit, Error> {
        let this = self;
        let package = this.package.clone();
        let patch = this.into_inner()?.finish()?;
        commit_change(&package, StoryTextChange(patch))
    }
}

impl Package {
    /// Stage a text change of one selected story.
    ///
    /// # Errors
    ///
    /// The selected story is admitted on the first operation, projection or
    /// commit, under the package's retained story limits.
    pub fn edit_story_text(&self, selector: Selector) -> Result<Edit, Error> {
        let limits = self.semantic_limits.story_text_value();
        Ok(Edit {
            package: self.clone(),
            selector,
            limits,
            inner: RefCell::new(None),
        })
    }

    /// Replay a selected-story patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] or [`Error::TopologyConflict`] for
    /// another story state.
    pub fn apply_story_text(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_change(self, patch)
    }
}
