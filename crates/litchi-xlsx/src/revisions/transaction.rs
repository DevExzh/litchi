//! Clone-staged semantic transactions for workbook revision metadata.

use litchi_opc::OpcPackage;

use crate::error::{Result, invalid};

use super::model::{RevisionHeader, RevisionLogPart, RevisionRecord, RevisionUser, Revisions};
use super::package::replace_workbook_revisions;
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::validation;

/// Failure-atomic edits over the workbook revision owner.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<Revisions>,
    conformance: super::model::RevisionConformance,
}

impl<'a> Transaction<'a> {
    /// Start a transaction from the package's main workbook.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        let draft = before.revisions().cloned();
        let conformance = before
            .conformance()
            .unwrap_or(super::model::RevisionConformance::Transitional);
        Ok(Self {
            target,
            before,
            draft,
            conformance,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged revision package.
    pub fn revisions(&self) -> Option<&Revisions> {
        self.draft.as_ref()
    }

    /// Replace or remove the complete revision package.
    pub fn replace(&mut self, value: Option<Revisions>) -> Result<bool> {
        if let Some(value) = &value {
            validation::revisions(value)?;
        }
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    /// Apply a checked mutation to the complete typed package.
    pub fn edit(&mut self, edit: impl FnOnce(&mut Revisions) -> Result<()>) -> Result<()> {
        let mut value = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        edit(&mut value)?;
        validation::revisions(&value)?;
        self.draft = Some(value);
        Ok(())
    }

    /// Insert or replace one typed revision user by its stable numeric ID.
    pub fn set_user(&mut self, value: RevisionUser) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit users on an absent revision package"))?;
        let Some(index) = draft
            .users
            .users
            .iter()
            .position(|user| user.id == value.id)
        else {
            draft.users.users.push(value);
            validation::revisions(&draft)?;
            self.draft = Some(draft);
            return Ok(true);
        };
        if draft.users.users[index] == value {
            return Ok(false);
        }
        draft.users.users[index] = value;
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Remove one typed revision user by stable numeric ID.
    pub fn remove_user(&mut self, id: i32) -> Result<Option<RevisionUser>> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit users on an absent revision package"))?;
        let Some(index) = draft.users.users.iter().position(|user| user.id == id) else {
            return Ok(None);
        };
        let removed = draft.users.users.remove(index);
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(Some(removed))
    }

    /// Insert or replace a header and its matching inert log as one unit.
    pub fn set_revision(&mut self, header: RevisionHeader, log: RevisionLogPart) -> Result<bool> {
        if header.relationship_id != log.relationship_id {
            return Err(invalid("revision header and log relationship IDs differ"));
        }
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        let header_index = draft
            .headers
            .headers
            .iter()
            .position(|candidate| candidate.relationship_id == header.relationship_id);
        let log_index = draft
            .logs
            .iter()
            .position(|candidate| candidate.relationship_id == log.relationship_id);
        match (header_index, log_index) {
            (Some(header_index), Some(log_index)) => {
                if draft.headers.headers[header_index] == header && draft.logs[log_index] == log {
                    return Ok(false);
                }
                draft.headers.headers[header_index] = header;
                draft.logs[log_index] = log;
            },
            (None, None) => {
                draft.headers.headers.push(header);
                draft.logs.push(log);
            },
            _ => return Err(invalid("revision header/log catalog is inconsistent")),
        }
        if let Some(last) = draft.headers.headers.last() {
            draft.headers.properties.guid = last.guid.clone();
        }
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Remove a header and its matching log by relationship ID.
    pub fn remove_revision(&mut self, relationship_id: &str) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        let Some(header_index) = draft
            .headers
            .headers
            .iter()
            .position(|header| header.relationship_id == relationship_id)
        else {
            return Ok(false);
        };
        let Some(log_index) = draft
            .logs
            .iter()
            .position(|log| log.relationship_id == relationship_id)
        else {
            return Err(invalid("revision header has no matching log"));
        };
        draft.headers.headers.remove(header_index);
        draft.logs.remove(log_index);
        if let Some(last) = draft.headers.headers.last() {
            draft.headers.properties.guid = last.guid.clone();
        }
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Replace one existing inert log without changing its header.
    pub fn set_log(&mut self, value: RevisionLogPart) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        let Some(index) = draft
            .logs
            .iter()
            .position(|log| log.relationship_id == value.relationship_id)
        else {
            return Err(invalid("cannot add a log without its revision header"));
        };
        if draft.logs[index] == value {
            return Ok(false);
        }
        draft.logs[index] = value;
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Insert or replace a record by its revision ID in one existing log.
    pub fn set_record(&mut self, relationship_id: &str, value: RevisionRecord) -> Result<bool> {
        let revision_id = value
            .revision_id
            .ok_or_else(|| invalid("revision record CRUD requires a revision ID"))?;
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        let log = draft
            .logs
            .iter_mut()
            .find(|log| log.relationship_id == relationship_id)
            .ok_or_else(|| invalid("revision log is absent"))?;
        if let Some(index) = log
            .log
            .records
            .iter()
            .position(|record| record.revision_id == Some(revision_id))
        {
            if log.log.records[index] == value {
                return Ok(false);
            }
            log.log.records[index] = value;
        } else {
            log.log.records.push(value);
        }
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(true)
    }

    /// Remove a record by revision ID from one existing log.
    pub fn remove_record(
        &mut self,
        relationship_id: &str,
        revision_id: u32,
    ) -> Result<Option<RevisionRecord>> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent revision package"))?;
        let log = draft
            .logs
            .iter_mut()
            .find(|log| log.relationship_id == relationship_id)
            .ok_or_else(|| invalid("revision log is absent"))?;
        let Some(index) = log
            .log
            .records
            .iter()
            .position(|record| record.revision_id == Some(revision_id))
        else {
            return Ok(None);
        };
        let removed = log.log.records.remove(index);
        validation::revisions(&draft)?;
        self.draft = Some(draft);
        Ok(Some(removed))
    }

    /// Whether staged semantics differ from the source semantics.
    pub fn is_changed(&self) -> bool {
        self.before.revisions() != self.draft.as_ref()
    }

    /// Validate and publish the staged owner atomically.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let mut candidate = self.target.clone();
        replace_workbook_revisions(&mut candidate, self.draft.as_ref(), self.conformance)?;
        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.revisions() != self.draft.as_ref() {
            return Err(invalid(
                "revision package publication changed the staged model",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}
