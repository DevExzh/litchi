//! Failure-atomic snapshots, edits, and reversible footnote-layout patches.

use std::sync::Arc;

use crate::error::{Error, Result};

use super::codec;
use super::model::Layout;
use super::validation;

/// An immutable, cheaply clonable `sectPr` snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Option<Layout>,
}

impl Snapshot {
    /// Parse and retain a bounded section-property snapshot.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read(&xml)?.value;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Return the authored section XML without copying it.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return the direct Word 2012 layout; absence remains observable.
    #[must_use]
    pub const fn layout(&self) -> Option<Layout> {
        self.value
    }

    /// Alias for callers reasoning in terms of the XML property name.
    #[must_use]
    pub const fn footnote_columns(&self) -> Option<Layout> {
        self.layout()
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

/// A section-property edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Option<Layout>,
}

impl Transaction {
    /// Return the projected layout in this transaction.
    #[must_use]
    pub const fn layout(&self) -> Option<Layout> {
        self.next
    }

    /// Set or remove the direct Word 2012 footnote layout.
    pub fn set_layout(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        validation::validate_layout(value)?;
        self.next = value;
        Ok(self)
    }

    /// Alias using the XML property vocabulary.
    pub fn set_footnote_columns(&mut self, value: Option<Layout>) -> Result<&mut Self> {
        self.set_layout(value)
    }

    /// Remove the direct footnote layout marker.
    #[must_use]
    pub fn clear(&mut self) -> &mut Self {
        self.next = None;
        self
    }

    /// Validate and publish the edit without changing the source snapshot.
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

/// A lineage-independent, preconditioned reversible layout patch.
#[derive(Debug, Clone, Copy, Eq, PartialEq)]
pub struct Patch {
    before: Option<Layout>,
    after: Option<Layout>,
}

impl Patch {
    /// Return the expected source state.
    #[must_use]
    pub const fn before(&self) -> Option<Layout> {
        self.before
    }

    /// Return the state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Option<Layout> {
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

    /// Apply the patch only when the target has its expected source state.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.value != self.before {
            return Err(Error::InvalidFormat(
                "footnote-columns patch source state does not match its precondition".into(),
            ));
        }
        let xml = codec::rewrite(source.xml_bytes(), self.after)?;
        Snapshot::from_xml(xml)
    }
}
