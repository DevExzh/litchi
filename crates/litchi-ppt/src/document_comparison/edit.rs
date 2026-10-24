//! Package-level document-comparison edits (record 0357).
//!
//! [`Package::edit_document_comparison`] stages edits of the inert
//! `PowerPoint` 10 review metadata of the live document, [`Edit::commit`]
//! publishes a new [`Package`] through one append-only user edit, and
//! [`Package::apply_document_comparison`] replays a [`Patch`] onto the exact
//! package it was made from. Dropping an [`Edit`] discards it. No
//! comparison, merge, accept or reject operation is performed.

use super::model::{DiffFlags, DiffTree10, Limits, Review, ReviewingToolbarStates};
use super::package::{Editor as RecordEdit, Snapshot as PackageSnapshot};
use super::{Change, Commit, Error, Patch, Result};
use crate::Package;
use crate::scoped::{self, Base, Failure, Lazy};

/// A staged document-comparison edit over one immutable [`Package`].
#[derive(Debug)]
pub struct Edit {
    base: Base,
    inner: Lazy<RecordEdit>,
}

impl Edit {
    fn new(package: &Package) -> Result<Self> {
        let base = Base::capture(package)?;
        Ok(Self {
            inner: Lazy::new(),
            base,
        })
    }

    fn initialize_inner(&self) -> std::result::Result<RecordEdit, Failure> {
        let snapshot =
            PackageSnapshot::from_bytes_with_limits(self.base.bytes()?.to_vec(), Limits::default())
                .map_err(scoped::publication("edit_document_comparison"))?;
        Ok(snapshot.edit())
    }

    fn inner(&self) -> std::result::Result<&RecordEdit, Failure> {
        self.inner.get_or_try_init(|| self.initialize_inner())
    }

    fn inner_mut(&mut self) -> std::result::Result<&mut RecordEdit, Failure> {
        let base = self.base.clone();
        self.inner.get_or_try_init_mut(|| {
            let snapshot =
                PackageSnapshot::from_bytes_with_limits(base.bytes()?.to_vec(), Limits::default())
                    .map_err(scoped::publication("edit_document_comparison"))?;
            Ok::<RecordEdit, Failure>(snapshot.edit())
        })
    }

    /// The staged review records and opaque records, in source order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the staged review payload cannot be
    /// read.
    pub fn review(&self) -> Result<Review> {
        Ok(self
            .inner()?
            .review()
            .map_err(scoped::publication("document_comparison.review"))?)
    }

    /// The staged reviewing-toolbar state, when present.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the staged review payload cannot be
    /// read.
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        Ok(self
            .inner()?
            .toolbar()
            .map_err(scoped::publication("document_comparison.toolbar"))?)
    }

    /// The staged reviewer tree at `index`, when present.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the staged review payload cannot be
    /// read.
    pub fn diff_tree(&self, index: usize) -> Result<Option<DiffTree10>> {
        Ok(self
            .inner()?
            .diff_tree(index)
            .map_err(scoped::publication("document_comparison.diff_tree"))?)
    }

    /// The staged semantic changes, in order.
    #[must_use = "handle the fallible document-comparison changes result"]
    pub fn changes(&self) -> Result<&[Change]> {
        Ok(self.inner()?.changes())
    }

    fn require_tree(&self, operation: &'static str, tree_index: usize) -> Result<()> {
        if self.diff_tree(tree_index)?.is_some() {
            Ok(())
        } else {
            Err(
                scoped::not_found(operation, "the document has no reviewer tree at this index")
                    .into(),
            )
        }
    }

    /// Stage a set or replacement of the reviewing-toolbar state. An equal
    /// state is an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the live document has no `___PPT10`
    /// review payload to hold the state, leaving the edit unchanged.
    pub fn set_toolbar(&mut self, value: ReviewingToolbarStates) -> Result<()> {
        Ok(self
            .inner_mut()?
            .set_toolbar(value)
            .map_err(target("document_comparison.set_toolbar"))?)
    }

    /// Stage removal of the reviewing-toolbar state, preserving every other
    /// `PowerPoint` 10 record; returns whether a state was present.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] when the staged review payload cannot be
    /// rewritten, leaving the edit unchanged.
    pub fn clear_toolbar(&mut self) -> Result<bool> {
        Ok(self
            .inner_mut()?
            .clear_toolbar()
            .map_err(target("document_comparison.clear_toolbar"))?)
    }

    /// Stage a replacement of one reviewer name, in source order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with
    /// [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound)
    /// when no reviewer tree has `tree_index`, and with
    /// [`RefusalKind::InvalidValue`](litchi_core::edit::RefusalKind::InvalidValue)
    /// when the name cannot be stored, leaving the edit unchanged.
    pub fn set_reviewer_name(&mut self, tree_index: usize, value: impl Into<String>) -> Result<()> {
        self.require_tree("document_comparison.set_reviewer_name", tree_index)?;
        Ok(self
            .inner_mut()?
            .set_reviewer_name(tree_index, value)
            .map_err(scoped::staging("document_comparison.set_reviewer_name"))?)
    }

    /// Stage a replacement of one reviewer tree's document-level display
    /// flags. This changes only what the legacy reviewing UI displays; it
    /// never applies, rejects or generates any presentation change.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] with
    /// [`RefusalKind::NotFound`](litchi_core::edit::RefusalKind::NotFound)
    /// when no reviewer tree has `tree_index`, and with
    /// [`RefusalKind::InvalidValue`](litchi_core::edit::RefusalKind::InvalidValue)
    /// when the flags cannot be stored, leaving the edit unchanged.
    pub fn set_document_flags(&mut self, tree_index: usize, value: DiffFlags) -> Result<()> {
        self.require_tree("document_comparison.set_document_flags", tree_index)?;
        Ok(self
            .inner_mut()?
            .set_document_flags(tree_index, value)
            .map_err(scoped::staging("document_comparison.set_document_flags"))?)
    }

    /// Publish the staged review metadata as a new package. An edit with no
    /// net change publishes the source package unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed write to a signed package,
    /// [`Error::Refused`] when the package holds storage the publisher cannot
    /// preserve, and [`Error::Package`] when the result cannot be
    /// republished.
    pub fn commit(self) -> Result<Commit> {
        let Self { base, inner } = self;
        inner
            .get_or_try_init(|| {
                let snapshot = PackageSnapshot::from_bytes_with_limits(
                    base.bytes()?.to_vec(),
                    Limits::default(),
                )
                .map_err(scoped::publication("edit_document_comparison"))?;
                Ok::<RecordEdit, Failure>(snapshot.edit())
            })
            .map_err(Error::from)?;
        let inner = inner.into_inner().ok_or_else(|| {
            Failure::Package(crate::package::Error::Corrupted(
                "PPT document-comparison edit state was not initialized".into(),
            ))
        })?;
        let after = match inner.commit() {
            Ok(commit) => commit.snapshot().bytes().to_vec(),
            // The publisher refuses the signature storages a changed write
            // would invalidate; report the signature, not the storage.
            Err(_) if base.is_signed()? => return Err(Error::Signed),
            Err(error) => {
                return Err(scoped::publication("document_comparison.commit")(error).into());
            },
        };
        Ok(super::published(base.publish(after)?))
    }
}

/// Map a toolbar staging failure: a document without a review payload cannot
/// carry the state.
fn target(operation: &'static str) -> impl FnOnce(crate::package::Error) -> Failure {
    move |error| match error {
        crate::package::Error::InvalidFormat(_) => scoped::unsupported_target(
            operation,
            "the live document has no PowerPoint 10 review payload",
        ),
        error => Failure::Package(error),
    }
}

impl Package {
    /// Start a document-comparison edit over this package.
    ///
    /// # Errors
    ///
    /// An encrypted presentation is refused immediately. A live document
    /// that cannot be read or losslessly represent review metadata is
    /// reported by the first projection, staging operation, or commit.
    pub fn edit_document_comparison(&self) -> Result<Edit> {
        Edit::new(self)
    }

    /// Apply a document-comparison patch to the exact package it was made
    /// from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source.
    pub fn apply_document_comparison(&self, patch: &Patch) -> Result<Commit> {
        super::apply_patch(self, patch)
    }
}
