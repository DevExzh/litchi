//! Snapshot editing for `RgxOcxInfo`.

use super::model::{OcxInfo, RgxOcxInfo};
use super::validation::{self, MAX_INFO_COUNT};
use crate::package::{Error as PackageError, Result};

/// One reversible semantic change made by an [`Editor`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Change {
    /// A record was inserted at `index`.
    Insert { index: usize, info: OcxInfo },
    /// A record was removed from `index`.
    Remove { index: usize, info: OcxInfo },
    /// A record was replaced in place.
    Replace {
        index: usize,
        before: OcxInfo,
        after: OcxInfo,
    },
}

/// The compact semantic patch produced by one OLE-control edit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    changes: Vec<Change>,
}

impl Patch {
    /// The ordered changes made by the editor.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Whether the edit made no semantic changes.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.changes.is_empty()
    }
}

/// A committed immutable OLE-control snapshot and its semantic patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: RgxOcxInfo,
    patch: Patch,
}

impl Commit {
    /// The edited immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &RgxOcxInfo {
        &self.snapshot
    }

    /// The compact patch from the source snapshot.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit and return its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (RgxOcxInfo, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A transactional editor over an immutable `RgxOcxInfo` snapshot.
#[derive(Debug, Clone)]
pub struct Editor {
    infos: Vec<OcxInfo>,
    changes: Vec<Change>,
}

impl Editor {
    /// Begin an edit from an immutable snapshot.
    #[must_use]
    pub fn new(source: &RgxOcxInfo) -> Self {
        Self {
            infos: source.infos().to_vec(),
            changes: Vec::new(),
        }
    }

    /// The current working records.
    #[must_use]
    pub fn infos(&self) -> &[OcxInfo] {
        &self.infos
    }

    /// Insert a validated record before `index`.
    pub fn insert(&mut self, index: usize, info: OcxInfo) -> Result<()> {
        if index > self.infos.len() {
            return Err(corrupted("OcxInfo insertion index is out of bounds"));
        }
        validation::info(&info)?;
        self.ensure_cookie_available(info.cookie(), None)?;
        if self.infos.len() >= MAX_INFO_COUNT {
            return Err(corrupted("RgxOcxInfo count exceeds the metadata limit"));
        }
        self.infos.insert(index, info);
        self.changes.push(Change::Insert { index, info });
        Ok(())
    }

    /// Append a validated record.
    pub fn push(&mut self, info: OcxInfo) -> Result<()> {
        self.insert(self.infos.len(), info)
    }

    /// Replace one record while preserving cookie uniqueness.
    pub fn replace(&mut self, index: usize, info: OcxInfo) -> Result<()> {
        let before = *self
            .infos
            .get(index)
            .ok_or_else(|| corrupted("OcxInfo replacement index is out of bounds"))?;
        validation::info(&info)?;
        self.ensure_cookie_available(info.cookie(), Some(index))?;
        if before != info {
            self.infos[index] = info;
            self.changes.push(Change::Replace {
                index,
                before,
                after: info,
            });
        }
        Ok(())
    }

    /// Remove one record and return it.
    pub fn remove(&mut self, index: usize) -> Result<OcxInfo> {
        if index >= self.infos.len() {
            return Err(corrupted("OcxInfo removal index is out of bounds"));
        }
        let info = self.infos.remove(index);
        self.changes.push(Change::Remove { index, info });
        Ok(info)
    }

    /// Commit the working records as a validated immutable snapshot.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = RgxOcxInfo::try_new(self.infos)?;
        Ok(Commit {
            snapshot,
            patch: Patch {
                changes: self.changes,
            },
        })
    }

    fn ensure_cookie_available(&self, cookie: u32, except: Option<usize>) -> Result<()> {
        if self
            .infos
            .iter()
            .enumerate()
            .any(|(index, info)| Some(index) != except && info.cookie() == cookie)
        {
            return Err(corrupted("OcxInfo dwCookie values must be unique"));
        }
        Ok(())
    }
}

impl RgxOcxInfo {
    /// Begin a transactional edit from this immutable snapshot.
    #[must_use]
    pub fn edit(&self) -> Editor {
        Editor::new(self)
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
