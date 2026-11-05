//! The plain-paragraph copy family of a source-backed package (record
//! 0355a): duplicate one plain main-story paragraph at another position.
//!
//! [`Package::edit_plain_paragraph_copy`](crate::source_backed::Package::edit_plain_paragraph_copy)
//! stages an [`Edit`] whose commit returns a new source-backed package;
//! writing it rewrites only the main-document part.

use std::cell::{RefCell, RefMut};

use litchi_core::Position;

use crate::source_backed::Package;
use crate::source_backed::deferred::{Change, source_family};
use crate::source_backed::paragraph_copy::raw;
use crate::source_backed::paragraph_copy_record as record;

/// The record patch the family publishes.
#[derive(Clone)]
pub(crate) struct CopyChange(record::Patch);

impl Change for CopyChange {
    type Error = record::Error;

    fn is_noop(&self) -> bool {
        self.0.is_noop()
    }

    fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    fn publish(&self, package: &Package) -> Result<Package, record::Error> {
        let (main, xml) = package.plain_paragraph_copy_target(&self.0)?;
        if self.is_noop() {
            return Ok(package.clone());
        }
        Ok(package.deferred_part_overlay(main, xml)?)
    }
}

source_family!("plain-paragraph copy", CopyChange, raw::Error);

pub use raw::{EffectReport, LimitKind};

impl Patch {
    /// Return the deterministic exact-copy effect.
    #[must_use]
    pub fn effect_report(&self) -> EffectReport {
        self.change.0.effect_report()
    }

    /// Encode the durable record patch under the caller policy.
    pub fn to_bytes(&self, limits: raw::Limits) -> Result<Vec<u8>, Error> {
        self.durable()?.to_bytes(limits)
    }

    /// The durable record patch of this change, bound to the complete
    /// artifact it applies to: for a commit's patch, the artifact of the
    /// package it was committed on; for its inverse, the committed
    /// package's output. Serialize it with `to_bytes(limits)` and replay it with
    /// `Patch::from` (record 0355a).
    ///
    /// # Errors
    ///
    /// Returns a source or resource error while fingerprinting the artifact.
    pub fn durable(&self) -> Result<raw::Patch, Error> {
        Ok(raw::Patch::from_record(match &self.restore {
            Some((before, _)) => self.change.0.clone().rebound(before.output_fingerprint()?),
            None => self.change.0.clone(),
        }))
    }
}

impl From<raw::Patch> for Patch {
    /// Replay a record patch onto a package whose main document holds its
    /// before bytes.
    fn from(patch: raw::Patch) -> Self {
        Self {
            change: CopyChange(patch.into_record()),
            restore: None,
        }
    }
}

/// A staged copy of one plain paragraph. Its resource errors carry a
/// [`LimitKind`] dimension. The source is never changed:
/// [`Self::commit`] returns a new package.
#[derive(Clone)]
pub struct Edit {
    package: Package,
    limits: raw::Limits,
    inner: RefCell<Option<record::Edit>>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Edit").finish_non_exhaustive()
    }
}

impl Edit {
    fn materialize(&self) -> Result<record::Edit, Error> {
        Ok(self
            .package
            .plain_paragraph_copy_snapshot_bounded(self.limits)?
            .edit())
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
        RefMut::filter_map(self.inner.borrow_mut(), Option::as_mut)
            .map_err(|_| Error::InvalidDurable)
    }

    fn into_inner(mut self) -> Result<record::Edit, Error> {
        if self.inner.get_mut().is_none() {
            let inner = self.materialize()?;
            *self.inner.get_mut() = Some(inner);
        }
        self.inner
            .into_inner()
            .ok_or_else(|| raw::Error::InvalidDurable)
    }

    /// The package the edit is staged over.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// The projected main story.
    pub fn projected(&self) -> Result<raw::Tree, Error> {
        Ok(raw::Tree::from_record(
            self.inner_mut()?.projected().clone(),
        ))
    }

    /// Copy source paragraph `source` before insertion slot `before`, both
    /// in source order.
    ///
    /// # Errors
    ///
    /// Returns a bounds, refusal or limit error, staging nothing.
    pub fn copy_plain_paragraph(
        &mut self,
        source: Position,
        before: Position,
    ) -> Result<&mut Self, Error> {
        self.inner_mut()?.copy_plain_paragraph(source, before)?;
        Ok(self)
    }

    /// The effect the commit will have.
    pub fn effect_report(&self) -> Result<EffectReport, Error> {
        Ok(self.inner_mut()?.clone().commit().effect_report())
    }

    /// Commit the copy as a new source-backed package.
    ///
    /// # Errors
    ///
    /// Returns a typed source, signature or limit error before any output.
    pub fn commit(self) -> Result<Commit, Error> {
        let this = self;
        let package = this.package.clone();
        let patch = this.into_inner()?.commit().patch().clone();
        commit_change(&package, CopyChange(patch))
    }
}

impl Package {
    /// Stage a plain-paragraph copy of the main story.
    ///
    /// # Errors
    ///
    /// The bounded plain-paragraph admission runs on the first projection,
    /// operation or commit.
    pub fn edit_plain_paragraph_copy(&self) -> Result<Edit, Error> {
        let limits = self.semantic_limits.paragraph_copy_value();
        Ok(Edit {
            package: self.clone(),
            limits,
            inner: RefCell::new(None),
        })
    }

    /// Replay a plain-paragraph copy patch onto this package.
    ///
    /// # Errors
    ///
    /// Returns a patch-conflict error for another main story.
    pub fn apply_plain_paragraph_copy(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_change(self, patch)
    }
}
