//! Immutable views and atomic record-tree edits for one master layout.

use super::codec;
use super::model::{Context, Inventory, Path, Snapshot};
use super::validation;
use crate::package::{Error, Result};
use crate::records::Record;

/// A deterministic content revision used by parent owners for conflict checks.
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

/// One reversible structural change made by a transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Change {
    /// A record was inserted into a parent container.
    Add {
        parent: Path,
        index: usize,
        record: Record,
    },
    /// A record was removed from its parent container.
    Remove { path: Path, record: Record },
    /// A record was replaced in place.
    Replace {
        path: Path,
        before: Record,
        after: Record,
    },
}

/// The changes and revisions produced by one successful commit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ChangeSet {
    base: Revision,
    target: Revision,
    changes: Vec<Change>,
}

impl ChangeSet {
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }

    /// Apply the inverse operations to the exact committed target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.target {
            return invalid("cannot undo against a different master-layout revision");
        }
        let mut edit = current.edit();
        for change in self.changes.iter().rev() {
            match change {
                Change::Add { parent, index, .. } => {
                    edit.remove(parent.child(*index))?;
                },
                Change::Remove { path, record } => {
                    let (parent, index) = split_path(path)?;
                    edit.add(parent, index, record.clone())?;
                },
                Change::Replace { path, before, .. } => {
                    edit.replace(path.clone(), before.clone())?;
                },
            }
        }
        Ok(edit.commit()?.snapshot().clone())
    }

    /// Reapply the operations to the exact committed source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.base {
            return invalid("cannot redo against a different master-layout revision");
        }
        let mut edit = current.edit();
        for change in &self.changes {
            match change {
                Change::Add {
                    parent,
                    index,
                    record,
                } => edit.add(parent.clone(), *index, record.clone())?,
                Change::Remove { path, .. } => {
                    edit.remove(path.clone())?;
                },
                Change::Replace { path, after, .. } => {
                    edit.replace(path.clone(), after.clone())?;
                },
            }
        }
        Ok(edit.commit()?.snapshot().clone())
    }
}

/// A successful commit containing the new immutable snapshot and its patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    changes: ChangeSet,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[must_use]
    pub const fn changes(&self) -> &ChangeSet {
        &self.changes
    }

    pub fn into_parts(self) -> (Snapshot, ChangeSet) {
        (self.snapshot, self.changes)
    }
}

/// An immutable borrowed view of a transaction-local record tree.
#[derive(Debug, Clone, Copy)]
pub struct View<'a> {
    context: Context,
    record: &'a Record,
    revision: Revision,
}

impl<'a> View<'a> {
    pub(super) const fn new(context: Context, record: &'a Record, revision: Revision) -> Self {
        Self {
            context,
            record,
            revision,
        }
    }

    #[must_use]
    pub const fn context(self) -> Context {
        self.context
    }

    #[must_use]
    pub const fn record(self) -> &'a Record {
        self.record
    }

    #[must_use]
    pub const fn revision(self) -> Revision {
        self.revision
    }

    pub fn inventory(self) -> Result<Inventory> {
        super::inventory::inventory(self.record)
    }
}

/// An isolated transactional editor. Every mutator works on a private clone;
/// commit encodes, reparses, and validates the candidate before publishing it.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    root: Record,
    changes: Vec<Change>,
}

impl Transaction {
    pub(super) fn open(source: Snapshot) -> Self {
        Self {
            root: source.root.clone(),
            changes: Vec::new(),
            source,
        }
    }

    #[must_use]
    pub const fn context(&self) -> Context {
        self.source.context
    }

    #[must_use]
    pub fn source(&self) -> &Snapshot {
        &self.source
    }

    #[must_use]
    pub fn view(&self) -> View<'_> {
        View::new(self.context(), &self.root, self.revision())
    }

    #[must_use]
    pub fn record(&self) -> &Record {
        &self.root
    }

    #[must_use]
    pub fn revision(&self) -> Revision {
        Revision::from_bytes(&self.encoded_for_revision())
    }

    #[must_use]
    pub fn is_dirty(&self) -> bool {
        !self.changes.is_empty()
    }

    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.is_dirty()
    }

    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Capture the current validated candidate as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        snapshot_from_root(self.context(), self.root.clone(), self.source.limits)
    }

    /// Insert one child at a checked position under `parent`.
    pub fn add(&mut self, parent: impl Into<Path>, index: usize, record: Record) -> Result<()> {
        let parent = parent.into();
        let limits = self.source.limits;
        let _ = codec::encode(&record, limits)?;
        let mut candidate = self.root.clone();
        let child_count = {
            let target = locate_mut(&mut candidate, parent.as_slice(), limits)?;
            codec::materialize(target, limits)?;
            target.children.len()
        };
        let _ = codec::encode(&candidate, limits)?;
        if index > child_count {
            return invalid("master-layout insertion index is out of range");
        }
        let target = locate_mut(&mut candidate, parent.as_slice(), limits)?;
        target.children.insert(index, record.clone());
        codec::sync(&mut candidate, limits)?;
        self.root = candidate;
        self.changes.push(Change::Add {
            parent,
            index,
            record,
        });
        Ok(())
    }

    /// Append one child to a container.
    pub fn append(&mut self, parent: impl Into<Path>, record: Record) -> Result<()> {
        let parent = parent.into();
        let index = child_count(&mut self.root, parent.as_slice(), self.source.limits)?;
        self.add(parent, index, record)
    }

    /// Remove one child. The source snapshot remains unchanged even if the
    /// resulting candidate later fails validation at commit.
    pub fn remove(&mut self, path: impl Into<Path>) -> Result<Record> {
        let path = path.into();
        let limits = self.source.limits;
        let mut candidate = self.root.clone();
        let removed = remove_at(&mut candidate, path.as_slice(), limits)?;
        codec::sync(&mut candidate, limits)?;
        self.root = candidate;
        self.changes.push(Change::Remove {
            path,
            record: removed.clone(),
        });
        Ok(removed)
    }

    /// Replace one record in place, including the root when `path` is empty.
    pub fn replace(&mut self, path: impl Into<Path>, record: Record) -> Result<Record> {
        let path = path.into();
        let limits = self.source.limits;
        let _ = codec::encode(&record, limits)?;
        let mut candidate = self.root.clone();
        let before = replace_at(&mut candidate, path.as_slice(), record.clone(), limits)?;
        codec::sync(&mut candidate, limits)?;
        self.root = candidate;
        self.changes.push(Change::Replace {
            path,
            before: before.clone(),
            after: record,
        });
        Ok(before)
    }

    /// Validate and publish a new immutable snapshot atomically.
    pub fn commit(self) -> Result<Commit> {
        let limits = self.source.limits;
        let bytes = codec::encode(&self.root, limits)?;
        let root = codec::parse(&bytes, limits)?;
        validation::validate(self.context(), &root, limits)?;
        let target = Snapshot {
            context: self.context(),
            root,
            bytes,
            limits,
        };
        let changes = ChangeSet {
            base: self.source.revision(),
            target: target.revision(),
            changes: self.changes,
        };
        Ok(Commit {
            snapshot: target,
            changes,
        })
    }

    /// Alias for callers using move-owned writer terminology.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard the private candidate and recover the source snapshot.
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn encoded_for_revision(&self) -> Vec<u8> {
        codec::encode(&self.root, self.source.limits).unwrap_or_default()
    }
}

fn child_count(root: &mut Record, path: &[usize], limits: super::model::Limits) -> Result<usize> {
    let count = {
        let parent = locate_mut(root, path, limits)?;
        codec::materialize(parent, limits)?;
        parent.children.len()
    };
    let _ = codec::encode(root, limits)?;
    Ok(count)
}

fn locate_mut<'a>(
    root: &'a mut Record,
    path: &[usize],
    limits: super::model::Limits,
) -> Result<&'a mut Record> {
    let mut current = root;
    for &index in path {
        codec::materialize(current, limits)?;
        current = current
            .children
            .get_mut(index)
            .ok_or_else(|| Error::InvalidFormat("master-layout path is out of range".into()))?;
    }
    Ok(current)
}

fn remove_at(root: &mut Record, path: &[usize], limits: super::model::Limits) -> Result<Record> {
    let (index, parent_path) = path
        .split_last()
        .ok_or_else(|| Error::InvalidFormat("the master-layout root cannot be removed".into()))?;
    {
        let parent = locate_mut(root, parent_path, limits)?;
        codec::materialize(parent, limits)?;
    }
    let _ = codec::encode(root, limits)?;
    let parent = locate_mut(root, parent_path, limits)?;
    if *index >= parent.children.len() {
        return invalid("master-layout removal path is out of range");
    }
    Ok(parent.children.remove(*index))
}

fn replace_at(
    root: &mut Record,
    path: &[usize],
    replacement: Record,
    limits: super::model::Limits,
) -> Result<Record> {
    if path.is_empty() {
        return Ok(std::mem::replace(root, replacement));
    }
    let Some((index, parent_path)) = path.split_last() else {
        return Err(Error::InvalidFormat(
            "the master-layout root cannot be replaced here".into(),
        ));
    };
    {
        let parent = locate_mut(root, parent_path, limits)?;
        codec::materialize(parent, limits)?;
    }
    let _ = codec::encode(root, limits)?;
    let parent = locate_mut(root, parent_path, limits)?;
    let slot = parent.children.get_mut(*index).ok_or_else(|| {
        Error::InvalidFormat("master-layout replacement path is out of range".into())
    })?;
    Ok(std::mem::replace(slot, replacement))
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

fn split_path(path: &Path) -> Result<(Path, usize)> {
    let (index, parent) = path
        .as_slice()
        .split_last()
        .ok_or_else(|| Error::InvalidFormat("the master-layout root cannot be addressed".into()))?;
    Ok((Path::from(parent), *index))
}

fn snapshot_from_root(
    context: Context,
    mut root: Record,
    limits: super::model::Limits,
) -> Result<Snapshot> {
    codec::sync(&mut root, limits)?;
    let bytes = codec::encode(&root, limits)?;
    let root = codec::parse(&bytes, limits)?;
    validation::validate(context, &root, limits)?;
    Ok(Snapshot {
        context,
        root,
        bytes,
        limits,
    })
}
