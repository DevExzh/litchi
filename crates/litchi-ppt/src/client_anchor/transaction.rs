//! Immutable snapshots and reversible anchor edits.

use std::sync::Arc;

use crate::package::{Error, Result};

use super::{Anchor, Data, Limits};

/// Deterministic identity of an exact serialized anchor record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn of(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// Immutable, cheap-to-clone owner of one exact anchor record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    anchor: Anchor,
    revision: Revision,
    limits: Limits,
}

impl Snapshot {
    /// Parse one serialized anchor record using the default bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or payload is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one serialized anchor record with an explicit resource bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or payload is malformed or exceeds `limits`.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let record = bytes.as_ref();
        let anchor = Anchor::parse_with_limits(record, limits)?;
        Ok(Self::from_parts(Arc::from(record), anchor, limits))
    }

    #[must_use]
    pub fn from_anchor(anchor: Anchor) -> Self {
        let bytes: Arc<[u8]> = Arc::from(anchor.to_bytes());
        Self::from_parts(bytes, anchor, Limits::default())
    }

    fn from_parts(bytes: Arc<[u8]>, anchor: Anchor, limits: Limits) -> Self {
        let revision = Revision::of(&bytes);
        Self {
            bytes,
            anchor,
            revision,
            limits,
        }
    }

    #[must_use]
    pub const fn anchor(&self) -> Anchor {
        self.anchor
    }

    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.anchor,
        }
    }
}

/// The single bounded semantic replacement represented by a patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Change {
    before: Anchor,
    after: Anchor,
}

impl Change {
    #[must_use]
    pub const fn before(self) -> Anchor {
        self.before
    }

    #[must_use]
    pub const fn after(self) -> Anchor {
        self.after
    }
}

/// Source-checked reversible replacement of one complete anchor record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    change: Option<Change>,
    limits: Limits,
}

impl Patch {
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    #[must_use]
    pub const fn change(&self) -> Option<Change> {
        self.change
    }

    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.change.is_none()
    }

    /// Apply this patch to its base snapshot, yielding the patched snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if `current` is not the patch base, or if the patched
    /// bytes fail to parse under the patch limits.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision != self.base || current.bytes.as_ref() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "anchor patch source does not match its base snapshot".into(),
            ));
        }
        Snapshot::parse_with_limits(self.after.as_ref(), self.limits)
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            change: self.change.map(|change| Change {
                before: change.after,
                after: change.before,
            }),
            limits: self.limits,
        }
    }
}

/// Isolated semantic edit over one source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Anchor,
}

impl Transaction {
    #[must_use]
    pub const fn anchor(&self) -> Anchor {
        self.candidate
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.anchor
    }

    pub fn set(&mut self, anchor: Anchor) {
        self.candidate = anchor;
    }

    pub fn set_data(&mut self, data: Data) {
        self.set(Anchor::new(data));
    }

    /// Replace the candidate with compact `SmallRectStruct` coordinates.
    ///
    /// # Errors
    ///
    /// Returns an error if `left` exceeds `right` or `top` exceeds `bottom`.
    pub fn set_small(&mut self, left: i16, top: i16, right: i16, bottom: i16) -> Result<()> {
        self.set(Anchor::small(left, top, right, bottom)?);
        Ok(())
    }

    /// Replace the candidate with full `RectStruct` coordinates.
    ///
    /// # Errors
    ///
    /// Returns an error if `left` exceeds `right` or `top` exceeds `bottom`.
    pub fn set_full(&mut self, left: i32, top: i32, right: i32, bottom: i32) -> Result<()> {
        self.set(Anchor::full(left, top, right, bottom)?);
        Ok(())
    }

    /// Materialize the candidate anchor as an immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the serialized candidate fails to re-parse under the
    /// source snapshot limits.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        Snapshot::parse_with_limits(self.candidate.to_bytes(), self.source.limits)
    }

    /// Publish the candidate as an immutable snapshot plus its reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the serialized candidate fails to re-parse under the
    /// source snapshot limits.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = if self.candidate == self.source.anchor {
            self.source.clone()
        } else {
            Snapshot::parse_with_limits(self.candidate.to_bytes(), self.source.limits)?
        };
        let change = (self.candidate != self.source.anchor).then_some(Change {
            before: self.source.anchor,
            after: self.candidate,
        });
        let patch = Patch {
            base: self.source.revision,
            target: snapshot.revision,
            before: self.source.bytes.clone(),
            after: snapshot.bytes.clone(),
            change,
            limits: self.source.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// Published immutable snapshot and its reversible patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[must_use]
    pub const fn anchor(&self) -> Anchor {
        self.snapshot.anchor
    }
}
