//! Immutable caption snapshots and failure-atomic semantic edits.

use super::model::{AutoTable, LabelTable};
use super::semantic::Tables;
use crate::package::{Error as PackageError, Result as PackageResult};

/// An immutable snapshot of the paired caption metadata tables.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    tables: Tables,
}

impl Snapshot {
    /// Creates a validated snapshot from caption tables.
    pub fn new(tables: Tables) -> PackageResult<Self> {
        let tables = Tables::try_new(tables.labels().cloned(), tables.auto().cloned())?;
        Ok(Self { tables })
    }

    /// Creates an empty snapshot with both FIB ranges absent.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            tables: Tables::empty(),
        }
    }

    /// Returns the paired caption metadata.
    #[must_use]
    pub fn tables(&self) -> &Tables {
        &self.tables
    }

    /// Returns the caption label table, when present.
    #[must_use]
    pub fn labels(&self) -> Option<&LabelTable> {
        self.tables.labels()
    }

    /// Returns the automatic-caption table, when present.
    #[must_use]
    pub fn auto(&self) -> Option<&AutoTable> {
        self.tables.auto()
    }

    /// Whether at least one caption range is active.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.labels().is_some() || self.auto().is_some()
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            working: self.clone(),
        }
    }
}

impl Tables {
    /// Creates an immutable snapshot of these caption tables.
    pub fn snapshot(&self) -> PackageResult<Snapshot> {
        Snapshot::new(self.clone())
    }

    /// Starts a validated semantic transaction from these caption tables.
    pub fn edit(&self) -> PackageResult<Transaction> {
        Ok(self.snapshot()?.edit())
    }
}

/// A staged, reversible caption metadata edit.
#[derive(Debug, Clone)]
pub struct Transaction {
    before: Snapshot,
    working: Snapshot,
}

impl Transaction {
    /// Returns the immutable source snapshot.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Returns the current staged snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Whether the staged semantic state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before != self.working
    }

    /// Replaces both caption tables as one validated operation.
    pub fn replace(&mut self, tables: Tables) -> Result<(), TransactionError> {
        self.working = Snapshot::new(tables).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Replaces the label definitions while retaining automatic-caption rules.
    pub fn replace_labels(&mut self, labels: LabelTable) -> Result<(), TransactionError> {
        let tables = Tables::try_new(Some(labels), self.working.auto().cloned())
            .map_err(TransactionError::Invalid)?;
        self.working = Snapshot::new(tables).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Replaces automatic-caption rules while retaining label definitions.
    pub fn replace_auto(&mut self, auto: AutoTable) -> Result<(), TransactionError> {
        let tables = Tables::try_new(self.working.labels().cloned(), Some(auto))
            .map_err(TransactionError::Invalid)?;
        self.working = Snapshot::new(tables).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Removes the label range. Non-empty rules that would dangle are rejected.
    pub fn clear_labels(&mut self) -> Result<(), TransactionError> {
        let tables = Tables::try_new(None, self.working.auto().cloned())
            .map_err(TransactionError::Invalid)?;
        self.working = Snapshot::new(tables).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Removes the automatic-caption range while retaining labels.
    pub fn clear_auto(&mut self) -> Result<(), TransactionError> {
        let tables = Tables::try_new(self.working.labels().cloned(), None)
            .map_err(TransactionError::Invalid)?;
        self.working = Snapshot::new(tables).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Removes both caption ranges.
    pub fn clear(&mut self) {
        self.working = Snapshot::empty();
    }

    /// Commits the staged semantic edit and returns its reversible patch.
    pub fn commit(self) -> Result<Commit, TransactionError> {
        Ok(Commit {
            before: self.before,
            after: self.working,
        })
    }
}

/// A reversible semantic change between two caption snapshots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    before: Snapshot,
    after: Snapshot,
}

impl Commit {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// The source snapshot.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The result snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the edit changed no semantic values.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies this patch to the exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, TransactionError> {
        if source != &self.before {
            return Err(TransactionError::Conflict);
        }
        Ok(self.after.clone())
    }

    /// Returns the exact inverse semantic patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Errors produced while staging or applying a caption edit.
#[derive(Debug)]
pub enum TransactionError {
    /// The candidate violates a bounded caption invariant.
    Invalid(PackageError),
    /// The transaction was created from a different snapshot.
    Conflict,
}

impl std::fmt::Display for TransactionError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("caption transaction snapshot conflict"),
        }
    }
}

impl std::error::Error for TransactionError {}
