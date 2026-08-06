//! Atomic semantic edits for one slide synchronization container.

use super::codec;
use super::model::{Limits, Snapshot, Synchronization};
use super::validation;
use crate::package::{Error, Result};
use crate::records::Record;

/// A deterministic content revision used for optimistic parent integration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    pub(super) fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Return the compact revision value for diagnostics or parent matching.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// One semantic replacement of the optional synchronization container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    index: usize,
    before: Option<Record>,
    after: Option<Record>,
}

impl Change {
    /// Root-child index affected by this change.
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Opaque source record, if one existed.
    pub const fn before(&self) -> Option<&Record> {
        self.before.as_ref()
    }

    /// Opaque committed record, if one exists.
    pub const fn after(&self) -> Option<&Record> {
        self.after.as_ref()
    }
}

/// Reusable patch and revision pair produced by one successful commit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangeSet {
    base: Revision,
    target: Revision,
    changes: Vec<Change>,
}

impl ChangeSet {
    pub const fn base(&self) -> Revision {
        self.base
    }

    pub const fn target(&self) -> Revision {
        self.target
    }

    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Undo this patch against the exact committed target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.target {
            return invalid("cannot undo against a different slide revision");
        }
        let mut editor = Editor::open(current.clone());
        for change in self.changes.iter().rev() {
            editor.apply_change(change, false)?;
        }
        editor.commit().map(|commit| commit.snapshot().clone())
    }

    /// Redo this patch against the exact source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.base {
            return invalid("cannot redo against a different slide revision");
        }
        let mut editor = Editor::open(current.clone());
        for change in &self.changes {
            editor.apply_change(change, true)?;
        }
        editor.commit().map(|commit| commit.snapshot().clone())
    }
}

/// An isolated transaction over one slide's synchronization metadata.
#[derive(Debug, Clone)]
pub struct Editor {
    source: Snapshot,
    root: Record,
    changes: Vec<Change>,
}

impl Editor {
    pub(super) fn open(source: Snapshot) -> Self {
        Self {
            root: source.root.clone(),
            source,
            changes: Vec::new(),
        }
    }

    /// Borrow the candidate slide record.
    pub const fn record(&self) -> &Record {
        &self.root
    }

    /// Read the candidate synchronization metadata.
    pub fn synchronization(&self) -> Result<Option<Synchronization>> {
        codec::read(&self.root)
    }

    /// Return the source limits used by this transaction.
    pub const fn limits(&self) -> Limits {
        self.source.limits
    }

    /// Whether a semantic change has been staged.
    pub const fn is_changed(&self) -> bool {
        !self.changes.is_empty()
    }

    /// Borrow the semantic changes staged so far.
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Set or replace synchronization metadata atomically in the candidate.
    pub fn set(&mut self, value: Synchronization) -> Result<()> {
        let replacement = codec::encode_sync(&value)?;
        let existing = validation::sync_index(&self.root)?;
        let (index, before) = match existing {
            Some(index) => {
                let before = self.root.children[index].clone();
                if before == replacement {
                    return Ok(());
                }
                (index, Some(before))
            },
            None => (validation::insertion_index(&self.root)?, None),
        };
        if before.is_some() {
            let slot = self.root.children.get_mut(index).ok_or_else(|| {
                Error::InvalidFormat(
                    "slide synchronization replacement index is out of range".into(),
                )
            })?;
            *slot = replacement.clone();
        } else if index <= self.root.children.len() {
            self.root.children.insert(index, replacement.clone());
        } else {
            return invalid("slide synchronization insertion index is out of range");
        }
        self.changes.push(Change {
            index,
            before,
            after: Some(replacement),
        });
        Ok(())
    }

    /// Remove synchronization metadata, returning whether it was present.
    pub fn clear(&mut self) -> Result<bool> {
        let Some(index) = validation::sync_index(&self.root)? else {
            return Ok(false);
        };
        let before = self.root.children.get(index).cloned();
        self.root.children.remove(index);
        self.changes.push(Change {
            index,
            before,
            after: None,
        });
        Ok(true)
    }

    /// Capture the candidate without publishing it.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::from_record_with_limits(self.root.clone(), self.source.limits)
    }

    /// Validate and publish the candidate and its reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = Snapshot::from_record_with_limits(self.root, self.source.limits)?;
        let changes = ChangeSet {
            base: self.source.revision(),
            target: snapshot.revision(),
            changes: self.changes,
        };
        Ok(Commit { snapshot, changes })
    }

    /// Alias for move-owned writer terminology.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard the candidate and recover the source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn apply_change(&mut self, change: &Change, forward: bool) -> Result<()> {
        let value = if forward {
            change.after.as_ref()
        } else {
            change.before.as_ref()
        };
        let was_absent = if forward {
            change.before.is_none()
        } else {
            change.after.is_none()
        };
        match (value, was_absent) {
            (Some(record), true) => {
                if change.index > self.root.children.len() {
                    return invalid("slide synchronization insertion index is out of range");
                }
                self.root.children.insert(change.index, record.clone());
            },
            (Some(record), false) => {
                let slot = self.root.children.get_mut(change.index).ok_or_else(|| {
                    Error::InvalidFormat(
                        "slide synchronization replacement index is out of range".into(),
                    )
                })?;
                *slot = record.clone();
            },
            (None, _) => {
                if change.index >= self.root.children.len() {
                    return invalid("slide synchronization removal index is out of range");
                }
                self.root.children.remove(change.index);
            },
        }
        Ok(())
    }
}

/// A successful immutable target and reusable patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    changes: ChangeSet,
}

impl Commit {
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    pub fn synchronization(&self) -> Result<Option<Synchronization>> {
        self.snapshot.synchronization()
    }

    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.changes.undo(current)
    }

    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.changes.redo(current)
    }

    pub fn into_parts(self) -> (Snapshot, ChangeSet) {
        (self.snapshot, self.changes)
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}
