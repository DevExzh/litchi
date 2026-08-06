//! Snapshot transaction for one fixed chart-area record.

use litchi_ograph::chart::Rect;

use crate::Result;

use super::model::{Change, Commit, Patch, Snapshot};
use super::validation;

/// A bounded editor that can replace only the existing chart-area geometry.
#[derive(Clone, Copy, Debug)]
pub struct Transaction {
    base: Snapshot,
    working: Snapshot,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Result<Self> {
        validation::ensure(source.rect())?;
        Ok(Self {
            base: source,
            working: source,
        })
    }

    /// Returns the current transaction snapshot.
    pub const fn snapshot(&self) -> Snapshot {
        self.working
    }

    /// Stages a replacement rectangle without changing record topology.
    pub fn set_rect(&mut self, rect: Rect) -> Result<&mut Self> {
        validation::ensure(rect)?;
        self.working = Snapshot::from_wire(rect);
        Ok(self)
    }

    /// Validates and publishes the source-checked change.
    pub fn commit(self) -> Result<Commit> {
        validation::ensure_pair(self.base.rect(), self.working.rect())?;
        let change = (self.base != self.working).then(|| Change::new(self.base, self.working));
        Ok(Commit::new(self.working, Patch::new(change)))
    }
}
