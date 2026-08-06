//! Failure-atomic semantic edits over one RTD stream snapshot.

use std::sync::Arc;

use crate::{Error, Result};

use super::super::model::{Cell, Record, UnknownRecord, Value};
use super::super::package::Package;
use super::{Commit, Patch, Snapshot};

/// A detached transaction over inert RTD topics, cached values, and subscriber
/// cells.  The source snapshot is never modified.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Vec<u8>,
    package: Package,
}

impl Transaction {
    pub(crate) fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.bytes().to_vec(),
            package: source.package().clone(),
            source,
        }
    }

    /// Borrow the immutable source snapshot used for publication checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Borrow the currently staged typed RTD records.
    #[must_use]
    pub fn records(&self) -> &[Record] {
        self.package.real_time_data()
    }

    /// Contextual alias for [`Self::records`].
    #[must_use]
    pub fn real_time_data(&self) -> &[Record] {
        self.records()
    }

    /// Contextual alias for [`Self::records`].
    #[must_use]
    pub fn topics(&self) -> &[Record] {
        self.records()
    }

    /// Borrow currently staged unknown records without copying their payloads.
    #[must_use]
    pub fn unknown_records(&self) -> impl Iterator<Item = UnknownRecord<'_>> + '_ {
        self.package.unknown_records(&self.candidate)
    }

    /// Materialize the staged candidate as a validated immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_slice() == self.source.bytes() {
            Ok(self.source.clone())
        } else {
            Snapshot::parse(&self.candidate)
        }
    }

    /// Whether any staged operation changes source bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.as_slice() != self.source.bytes()
    }

    /// Replace one typed RTD record by semantic collection index.
    pub fn replace(&mut self, index: usize, value: Record) -> Result<&mut Self> {
        let current = self.records().get(index).ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "RealTimeData index {index} is outside the staged collection"
            ))
        })?;
        if current == &value {
            return Ok(self);
        }

        let candidate = self
            .package
            .replace_real_time_data(&self.candidate, index, &value)?;
        self.replace_candidate(candidate, Some((index, value)))?;
        Ok(self)
    }

    /// Alias for [`Self::replace`] using the contextual record name.
    pub fn replace_record(&mut self, index: usize, value: Record) -> Result<&mut Self> {
        self.replace(index, value)
    }

    /// Mutate one typed RTD record through a failure-atomic closure.
    pub fn edit(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Record) -> Result<()>,
    ) -> Result<&mut Self> {
        let mut value = self.records().get(index).cloned().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "RealTimeData index {index} is outside the staged collection"
            ))
        })?;
        edit(&mut value)?;
        self.replace(index, value)
    }

    /// Replace only one topic's inert cached value.
    pub fn set_value(&mut self, index: usize, value: Value) -> Result<&mut Self> {
        self.edit(index, |record| {
            record.value = value;
            Ok(())
        })
    }

    /// Replace only one topic's subscriber cells.
    pub fn set_cells(&mut self, index: usize, cells: Vec<Cell>) -> Result<&mut Self> {
        self.edit(index, |record| {
            record.cells = cells;
            Ok(())
        })
    }

    /// Insert a typed RTD record before the staged collection index.
    pub fn insert(&mut self, index: usize, value: Record) -> Result<&mut Self> {
        let candidate = self
            .package
            .insert_real_time_data(&self.candidate, index, &value)?;
        self.replace_candidate(candidate, Some((index, value)))?;
        Ok(self)
    }

    /// Append a typed RTD record after the complete source stream.
    pub fn push(&mut self, value: Record) -> Result<&mut Self> {
        let index = self.records().len();
        self.insert(index, value)
    }

    /// Remove one typed RTD record and return its previous semantic value.
    pub fn remove(&mut self, index: usize) -> Result<Record> {
        let removed = self.records().get(index).cloned().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "RealTimeData index {index} is outside the staged collection"
            ))
        })?;
        let candidate = self.package.remove_real_time_data(&self.candidate, index)?;
        self.replace_candidate(candidate, None)?;
        Ok(removed)
    }

    /// Discard all staged edits and return the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validate and publish the staged candidate with a reversible patch.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source;
        if self.candidate.as_slice() == source.bytes() {
            let patch = Patch::new(source.clone(), source.clone());
            return Ok(Commit::new(source, patch));
        }
        let snapshot = Snapshot::parse_shared(Arc::from(self.candidate.into_boxed_slice()))?;
        let patch = Patch::new(source, snapshot.clone());
        Ok(Commit::new(snapshot, patch))
    }

    fn replace_candidate(
        &mut self,
        candidate: Vec<u8>,
        expected: Option<(usize, Record)>,
    ) -> Result<()> {
        // Parse before publishing the candidate.  This validates all prefix
        // dependencies, continuation framing, and bounds atomically.
        let package = Package::parse(&candidate)?;
        if let Some((index, expected)) = expected {
            if package.real_time_data().get(index) != Some(&expected) {
                return Err(Error::UnsafeEdit(
                    "RealTimeData publication changed the staged record".to_string(),
                ));
            }
        }
        self.candidate = candidate;
        self.package = package;
        Ok(())
    }
}
