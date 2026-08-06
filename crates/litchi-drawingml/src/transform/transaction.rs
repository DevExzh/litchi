//! Immutable transform snapshots and detached, failure-atomic edits.

use std::sync::Arc;

use crate::{Error, Result};

use super::{Angle, Point, Size, Transform, codec, validation};

/// An immutable, cheaply clonable source snapshot of one `a:xfrm` element.
///
/// The original bytes are replayed for a no-op commit, retaining producer
/// prefixes, attribute order, whitespace, and lexical defaults exactly.
/// Changed commits are rebuilt from the fully typed transform model.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Transform,
}

impl Snapshot {
    /// Parse and retain a bounded transform fragment.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Create a canonical snapshot from a detached transform value.
    pub fn new(value: Transform) -> Result<Self> {
        Self::from_xml(codec::write(&value)?)
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed transform value.
    #[inline]
    #[must_use]
    pub const fn value(&self) -> &Transform {
        &self.value
    }

    /// Alias using the shared DrawingML vocabulary.
    #[inline]
    #[must_use]
    pub const fn transform(&self) -> &Transform {
        self.value()
    }

    /// Start an isolated edit based on this immutable snapshot.
    #[inline]
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            working: self.value.clone(),
        }
    }
}

/// A transform edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    working: Transform,
}

impl Transaction {
    /// Borrow the projected typed transform.
    #[inline]
    #[must_use]
    pub const fn transform(&self) -> &Transform {
        &self.working
    }

    /// Replace the complete typed transform after validating its output
    /// budget.
    pub fn set(&mut self, value: Transform) -> Result<&mut Self> {
        validation::validate(&value)?;
        self.working = value;
        Ok(self)
    }

    /// Set or clear the object offset.
    #[inline]
    pub fn set_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.working.set_offset(value);
        self
    }

    /// Set or clear the object extent.
    #[inline]
    pub fn set_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.working.set_extent(value);
        self
    }

    /// Set or clear the group child-coordinate offset.
    #[inline]
    pub fn set_child_offset(&mut self, value: Option<Point>) -> &mut Self {
        self.working.set_child_offset(value);
        self
    }

    /// Set or clear the group child-coordinate extent.
    #[inline]
    pub fn set_child_extent(&mut self, value: Option<Size>) -> &mut Self {
        self.working.set_child_extent(value);
        self
    }

    /// Set or clear the authored rotation.
    #[inline]
    pub fn set_rotation(&mut self, value: Option<Angle>) -> &mut Self {
        self.working.set_rotation(value);
        self
    }

    /// Set or clear the authored horizontal flip.
    #[inline]
    pub fn set_flip_horizontal(&mut self, value: Option<bool>) -> &mut Self {
        self.working.set_flip_horizontal(value);
        self
    }

    /// Set or clear the authored vertical flip.
    #[inline]
    pub fn set_flip_vertical(&mut self, value: Option<bool>) -> &mut Self {
        self.working.set_flip_vertical(value);
        self
    }

    /// Whether this edit changes the typed snapshot.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.base.value != self.working
    }

    /// Validate and publish the edit without changing the source snapshot.
    pub fn commit(self) -> Result<Commit> {
        validation::validate(&self.working)?;
        let xml = if self.base.value == self.working {
            self.base.xml.as_ref().to_vec()
        } else {
            codec::write(&self.working)?
        };
        let snapshot = Snapshot::from_xml(xml)?;
        Ok(Commit {
            patch: Patch {
                before: self.base.value,
                after: self.working,
            },
            snapshot,
        })
    }
}

/// A successful transform publication containing a new snapshot and a
/// reversible semantic patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of this commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of this commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A lineage-independent, preconditioned reversible transform patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Transform,
    after: Transform,
}

impl Patch {
    /// Borrow the expected source state.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Transform {
        &self.before
    }

    /// Borrow the state produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Transform {
        &self.after
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(self) -> Self {
        Self {
            before: self.after,
            after: self.before,
        }
    }

    /// Apply the patch only when the target has the expected semantic state.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.value != self.before {
            return Err(Error::Invalid(
                "DrawingML transform patch source state does not match its precondition".into(),
            ));
        }
        let xml = if self.before == self.after {
            source.xml.as_ref().to_vec()
        } else {
            codec::write(&self.after)?
        };
        Snapshot::from_xml(xml)
    }
}
