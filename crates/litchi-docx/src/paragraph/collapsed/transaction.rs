#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Failure-atomic snapshots, edits, commits, and reversible patches.

use std::sync::Arc;

use crate::error::{Error, Result};

use super::codec;
use super::model::Collapsed;
use super::validation;

/// An immutable, cheaply clonable paragraph snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Option<Collapsed>,
}

impl Snapshot {
    /// Parse and retain a bounded paragraph snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Return the current authored paragraph XML without copying it.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return the direct `collapsed` state; absence remains observable.
    #[must_use]
    pub const fn collapsed(&self) -> Option<Collapsed> {
        self.value
    }

    /// Start an isolated edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.value,
        }
    }
}

/// A paragraph edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Option<Collapsed>,
}

impl Transaction {
    /// Return the projected value in this transaction.
    #[must_use]
    pub const fn collapsed(&self) -> Option<Collapsed> {
        self.next
    }

    /// Set or remove the direct Word 2012 collapse marker.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_collapsed(&mut self, value: Option<Collapsed>) -> Result<&mut Self> {
        validation::validate(value)?;
        self.next = value;
        Ok(self)
    }

    /// Remove the direct Word 2012 collapse marker.
    pub fn clear_collapsed(&mut self) -> &mut Self {
        self.next = None;
        self
    }

    /// Validate and publish the edit without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        let xml = codec::rewrite(self.base.xml_bytes(), self.next)?;
        let snapshot = Snapshot::from_xml(xml)?;
        Ok(Commit {
            patch: Patch {
                before: self.base.value,
                after: self.next,
            },
            snapshot,
        })
    }
}

/// A successful publication containing the new snapshot and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A one-property, lineage-independent reversible paragraph patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Patch {
    before: Option<Collapsed>,
    after: Option<Collapsed>,
}

impl Patch {
    /// Return the expected source state.
    #[must_use]
    pub const fn before(&self) -> Option<Collapsed> {
        self.before
    }

    /// Return the state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Option<Collapsed> {
        self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub const fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }

    /// Apply the patch only when the target has the expected source state.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.value != self.before {
            return Err(Error::InvalidFormat(
                "collapsed patch source state does not match its precondition".to_owned(),
            ));
        }
        let xml = codec::rewrite(source.xml_bytes(), self.after)?;
        Snapshot::from_xml(xml)
    }
}
