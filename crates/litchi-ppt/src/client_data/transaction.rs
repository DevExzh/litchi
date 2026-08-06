//! Source-preserving semantic edits for one OfficeArtClientData record.

use std::sync::Arc;

use crate::package::{Error, Result};

use super::model::{ClientData, ClientDataChild, ClientDataLimits};

/// Deterministic identity of one exact serialized client-data record.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Compact revision value useful for parent-owner conflict checks.
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// One ordered child edit made by a transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Change {
    /// Insert an ordered child before the child currently at `index`.
    Insert {
        index: usize,
        child: ClientDataChild,
    },
    /// Remove the child currently at `index`.
    Remove {
        index: usize,
        child: ClientDataChild,
    },
    /// Replace one child in place.
    Replace {
        index: usize,
        before: ClientDataChild,
        after: ClientDataChild,
    },
}

impl Change {
    /// Child-list index affected by this edit.
    pub const fn index(&self) -> usize {
        match self {
            Self::Insert { index, .. }
            | Self::Remove { index, .. }
            | Self::Replace { index, .. } => *index,
        }
    }

    /// The child present before this edit, when there was one.
    pub fn before(&self) -> Option<&ClientDataChild> {
        match self {
            Self::Insert { .. } => None,
            Self::Remove { child, .. } => Some(child),
            Self::Replace { before, .. } => Some(before),
        }
    }

    /// The child present after this edit, when there is one.
    pub fn after(&self) -> Option<&ClientDataChild> {
        match self {
            Self::Insert { child, .. } => Some(child),
            Self::Remove { .. } => None,
            Self::Replace { after, .. } => Some(after),
        }
    }
}

/// Immutable, cheap-to-clone source or committed client-data state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    value: ClientData,
    revision: Revision,
    limits: ClientDataLimits,
}

impl Snapshot {
    /// Parse one complete record and retain its exact source bytes.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, ClientDataLimits::default())
    }

    /// Parse one complete record under explicit resource bounds.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: ClientDataLimits) -> Result<Self> {
        let bytes = bytes.as_ref();
        let value = ClientData::parse_with_limits(bytes, limits)?;
        let encoded = value.to_bytes_with_limits(limits)?;
        if encoded != bytes {
            return Err(Error::Corrupted(
                "OfficeArtClientData is not losslessly representable".into(),
            ));
        }
        Ok(Self::from_parts(
            Arc::from(bytes.to_vec().into_boxed_slice()),
            value,
            limits,
        ))
    }

    /// Capture a validated semantic value using the default resource bounds.
    pub fn from_client_data(value: ClientData) -> Result<Self> {
        Self::from_client_data_with_limits(value, ClientDataLimits::default())
    }

    /// Capture a validated semantic value under explicit resource bounds.
    pub fn from_client_data_with_limits(
        value: ClientData,
        limits: ClientDataLimits,
    ) -> Result<Self> {
        let bytes = value.to_bytes_with_limits(limits)?;
        Self::parse_with_limits(bytes, limits)
    }

    fn from_parts(bytes: Arc<[u8]>, value: ClientData, limits: ClientDataLimits) -> Self {
        let revision = Revision::from_bytes(&bytes);
        Self {
            bytes,
            value,
            revision,
            limits,
        }
    }

    /// Borrow the validated semantic client-data value.
    pub const fn client_data(&self) -> &ClientData {
        &self.value
    }

    /// Borrow ordered typed and opaque children.
    pub fn children(&self) -> &[ClientDataChild] {
        self.value.children()
    }

    /// Exact source or committed record bytes.
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Revision of the exact serialized source.
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Resource bounds retained for subsequent edits and patch application.
    pub const fn limits(&self) -> ClientDataLimits {
        self.limits
    }

    /// Start an isolated semantic edit over this snapshot.
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.value.clone(),
            changes: Vec::new(),
        }
    }
}

/// Reversible, source-checked result of one successful transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    changes: Vec<Change>,
    limits: ClientDataLimits,
}

impl Patch {
    /// Source revision required by [`Self::apply`].
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Target revision produced by [`Self::apply`].
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Ordered semantic edits represented by this patch.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Whether the transaction changed no child bytes.
    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Apply only to the exact source snapshot from which this patch came.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision != self.base || current.bytes.as_ref() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "client-data patch source does not match its base snapshot".into(),
            ));
        }
        Snapshot::parse_with_limits(self.after.as_ref(), self.limits)
    }

    /// Apply the inverse to the exact committed target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(current)
    }

    /// Reapply this patch to its exact source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.apply(current)
    }

    /// Build a source-checked inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
            limits: self.limits,
        }
    }
}

impl Change {
    fn inverse(&self) -> Self {
        match self {
            Self::Insert { index, child } => Self::Remove {
                index: *index,
                child: child.clone(),
            },
            Self::Remove { index, child } => Self::Insert {
                index: *index,
                child: child.clone(),
            },
            Self::Replace {
                index,
                before,
                after,
            } => Self::Replace {
                index: *index,
                before: after.clone(),
                after: before.clone(),
            },
        }
    }
}

/// Isolated, failure-atomic semantic editor over one source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: ClientData,
    changes: Vec<Change>,
}

impl Transaction {
    /// Borrow the immutable source snapshot.
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Borrow the current transaction candidate.
    pub const fn client_data(&self) -> &ClientData {
        &self.candidate
    }

    /// Borrow the candidate's ordered child list.
    pub fn children(&self) -> &[ClientDataChild] {
        self.candidate.children()
    }

    /// Whether any staged operation changes the serialized candidate.
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.value
    }

    /// Borrow the staged semantic operations.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Insert one validated child at an exact ordered position.
    pub fn insert(&mut self, index: usize, child: ClientDataChild) -> Result<()> {
        if index > self.candidate.children.len() {
            return invalid("client-data insertion index is out of range");
        }
        let mut candidate = self.candidate.clone();
        candidate.children.insert(index, child.clone());
        validate_candidate(&candidate, self.source.limits)?;
        self.candidate = candidate;
        self.changes.push(Change::Insert { index, child });
        Ok(())
    }

    /// Append one validated child after the current ordered sequence.
    pub fn append(&mut self, child: ClientDataChild) -> Result<()> {
        self.insert(self.candidate.children.len(), child)
    }

    /// Replace one child in place, preserving all other child order.
    pub fn replace(
        &mut self,
        index: usize,
        replacement: ClientDataChild,
    ) -> Result<ClientDataChild> {
        let before = self.candidate.children.get(index).cloned().ok_or_else(|| {
            Error::InvalidFormat("client-data replacement index is out of range".into())
        })?;
        if before == replacement {
            return Ok(before);
        }
        let mut candidate = self.candidate.clone();
        candidate.children[index] = replacement.clone();
        validate_candidate(&candidate, self.source.limits)?;
        self.candidate = candidate;
        self.changes.push(Change::Replace {
            index,
            before: before.clone(),
            after: replacement,
        });
        Ok(before)
    }

    /// Remove one child and return the removed value.
    pub fn remove(&mut self, index: usize) -> Result<ClientDataChild> {
        if index >= self.candidate.children.len() {
            return Err(Error::InvalidFormat(
                "client-data removal index is out of range".into(),
            ));
        }
        let mut candidate = self.candidate.clone();
        let child = candidate.children.remove(index);
        validate_candidate(&candidate, self.source.limits)?;
        self.candidate = candidate;
        self.changes.push(Change::Remove {
            index,
            child: child.clone(),
        });
        Ok(child)
    }

    /// Capture the current candidate without publishing the transaction.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.candidate.to_bytes_with_limits(self.source.limits)?;
        if bytes.as_slice() == self.source.bytes.as_ref() {
            return Ok(self.source.clone());
        }
        Snapshot::parse_with_limits(bytes, self.source.limits)
    }

    /// Validate and publish the candidate atomically with its reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let bytes = self.candidate.to_bytes_with_limits(self.source.limits)?;
        let snapshot = if bytes.as_slice() == self.source.bytes.as_ref() {
            self.source.clone()
        } else {
            Snapshot::parse_with_limits(bytes, self.source.limits)?
        };
        let changes = if snapshot.revision == self.source.revision {
            Vec::new()
        } else {
            self.changes
        };
        let patch = Patch {
            base: self.source.revision,
            target: snapshot.revision,
            before: self.source.bytes.clone(),
            after: snapshot.bytes.clone(),
            changes,
            limits: self.source.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Alias for move-owned writer terminology.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard all staged edits and recover the source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

fn validate_candidate(candidate: &ClientData, limits: ClientDataLimits) -> Result<()> {
    candidate.validate_with_limits(limits)
}

/// A successful immutable target and its source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published target snapshot.
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from the source to this target.
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Published target semantic value.
    pub const fn client_data(&self) -> &ClientData {
        self.snapshot.client_data()
    }

    /// Split the published target and patch.
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
