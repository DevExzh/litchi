//! Atomic snapshot edits.

use crate::{Error, Result};

use super::{RECORD_TYPE, Record, Settings, Snapshot, Unknown};

/// A compare-and-apply snapshot patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    pub fn apply(&self, target: &mut Snapshot) -> Result<()> {
        if target != &self.before {
            return Err(Error::UnsafeEdit(
                "CompressPictures snapshot changed since the patch was prepared".to_string(),
            ));
        }
        self.after.validate()?;
        *target = self.after.clone();
        Ok(())
    }
}

/// Detached editor; the source snapshot is never modified.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    working: Snapshot,
}

impl Transaction {
    pub(crate) fn new(snapshot: Snapshot) -> Self {
        Self {
            base: snapshot.clone(),
            working: snapshot,
        }
    }

    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    pub fn set_settings(&mut self, value: Settings) -> Result<()> {
        let mut candidate = self.working.clone();
        if let Some(record) = candidate
            .records
            .iter_mut()
            .find(|record| matches!(record, Record::Settings(_)))
        {
            *record = Record::Settings(value);
        } else {
            candidate.records.push(Record::Settings(value));
        }
        candidate.validate()?;
        self.working = candidate;
        Ok(())
    }

    pub fn remove_settings(&mut self) -> Option<Settings> {
        let position = self
            .working
            .records
            .iter()
            .position(|record| matches!(record, Record::Settings(_)))?;
        match self.working.records.remove(position) {
            Record::Settings(value) => Some(value),
            Record::Unknown(_) => unreachable!(),
        }
    }

    pub fn insert_unknown(
        &mut self,
        index: usize,
        record_type: u16,
        payload: impl Into<Vec<u8>>,
    ) -> Result<()> {
        self.insert(
            index,
            Record::Unknown(Unknown::try_new(record_type, payload)?),
        )
    }

    pub fn insert(&mut self, index: usize, record: Record) -> Result<()> {
        if index > self.working.records.len() {
            return Err(invalid("record insertion index is out of bounds"));
        }
        let mut candidate = self.working.clone();
        candidate.records.insert(index, record);
        candidate.validate()?;
        self.working = candidate;
        Ok(())
    }

    pub fn remove(&mut self, index: usize) -> Result<Record> {
        let mut candidate = self.working.clone();
        if index >= candidate.records.len() {
            return Err(invalid("record removal index is out of bounds"));
        }
        let removed = candidate.records.remove(index);
        candidate.validate()?;
        self.working = candidate;
        Ok(removed)
    }

    pub fn commit(self) -> Result<Patch> {
        self.working.validate()?;
        Ok(Patch {
            before: self.base,
            after: self.working,
        })
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: RECORD_TYPE,
        message: message.into(),
    }
}
