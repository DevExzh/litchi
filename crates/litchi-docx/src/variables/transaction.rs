//! Source-checked snapshots, edits, commits, and patches for document variables.

use std::sync::Arc;

use crate::{Error, Result};

use super::codec;
use super::model::Variables;

/// An immutable settings XML snapshot with its typed document-variable view.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    variables: Variables,
}

impl Snapshot {
    /// Parse and retain a bounded complete Word settings XML snapshot.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let variables = Variables::from_xml(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            variables,
        })
    }

    /// Borrow the exact settings XML retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed document-variable collection in source order.
    #[inline]
    #[must_use]
    pub const fn variables(&self) -> &Variables {
        &self.variables
    }

    /// Alias for [`Self::variables`] when the owner is used as a collection.
    #[inline]
    #[must_use]
    pub const fn collection(&self) -> &Variables {
        self.variables()
    }

    /// Start an isolated source-checked edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.variables.clone(),
        }
    }

    fn same_source(&self, other: &Self) -> bool {
        self.xml.as_ref() == other.xml.as_ref() && self.variables == other.variables
    }
}

/// A document-variable edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Variables,
}

impl Transaction {
    /// Borrow the projected document-variable collection.
    #[inline]
    #[must_use]
    pub const fn variables(&self) -> &Variables {
        &self.next
    }

    /// Alias for [`Self::variables`].
    #[inline]
    #[must_use]
    pub const fn collection(&self) -> &Variables {
        self.variables()
    }

    /// Whether the staged collection differs from its source snapshot.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.next != self.base.variables
    }

    /// Insert or replace one document variable atomically.
    pub fn set(&mut self, name: impl Into<String>, value: impl Into<String>) -> Result<&mut Self> {
        self.next.insert(name, value)?;
        Ok(self)
    }

    /// Alias for [`Self::set`].
    pub fn insert(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        self.set(name, value)
    }

    /// Alias for [`Self::set`] with an owner-specific verb.
    pub fn set_variable(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        self.set(name, value)
    }

    /// Remove one variable and return its previous value.
    pub fn remove(&mut self, name: &str) -> Option<String> {
        self.next.remove(name)
    }

    /// Alias for [`Self::remove`].
    pub fn remove_variable(&mut self, name: &str) -> Option<String> {
        self.remove(name)
    }

    /// Remove every staged variable.
    pub fn clear(&mut self) -> &mut Self {
        self.next.clear();
        self
    }

    /// Alias for [`Self::clear`].
    pub fn clear_variables(&mut self) -> &mut Self {
        self.clear()
    }

    /// Replace the complete collection after validating all Word limits.
    pub fn replace(&mut self, value: Variables) -> Result<&mut Self> {
        value.validate()?;
        self.next = value;
        Ok(self)
    }

    /// Apply a clone-staged collection mutation atomically.
    pub fn edit_variables(
        &mut self,
        edit: impl FnOnce(&mut Variables) -> Result<()>,
    ) -> Result<&mut Self> {
        let mut candidate = self.next.clone();
        edit(&mut candidate)?;
        candidate.validate()?;
        self.next = candidate;
        Ok(self)
    }

    /// Alias for [`Self::edit_variables`].
    pub fn edit(&mut self, edit: impl FnOnce(&mut Variables) -> Result<()>) -> Result<&mut Self> {
        self.edit_variables(edit)
    }

    /// Validate and publish this edit without mutating its source snapshot.
    pub fn commit(self) -> Result<Commit> {
        self.next.validate()?;
        if self.next == self.base.variables {
            let patch = Patch {
                before: self.base.clone(),
                after: self.base.clone(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
                changed: false,
            });
        }

        let xml = codec::rewrite(self.base.xml_bytes(), &self.base.variables, &self.next)?;
        let snapshot = Snapshot::from_xml(xml)?;
        if snapshot.variables != self.next {
            return Err(invalid(
                "document-variable publication changed the staged collection",
            ));
        }
        let patch = Patch {
            before: self.base,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful document-variable publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the transaction changed the authored source.
    #[inline]
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the published snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Move both publication products out of the commit.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible patch with exact source and typed semantic preconditions.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Borrow the expected source collection.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Variables {
        self.before.variables()
    }

    /// Borrow the collection produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Variables {
        self.after.variables()
    }

    /// Borrow the exact source snapshot used by this patch.
    #[inline]
    #[must_use]
    pub const fn before_snapshot(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the exact snapshot produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after_snapshot(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch leaves both source semantics and bytes unchanged.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply only to a snapshot matching both source preconditions.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(Error::InvalidFormat(
                "document-variable patch source does not match its precondition".into(),
            ));
        }
        Ok(self.after.clone())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
