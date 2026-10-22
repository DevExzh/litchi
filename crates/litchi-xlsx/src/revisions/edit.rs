//! Clone-staged semantic transactions for workbook revision metadata.

use litchi_opc::OpcPackage;

use crate::error::{Result, invalid};

use super::model::{RevisionHeader, RevisionLogPart, RevisionRecord, RevisionUser, Revisions};
use super::package::replace_workbook_revisions;
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::Snapshot;
use super::validation;

/// The configuration of a revision edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    conformance: Option<super::model::RevisionConformance>,
}

impl EditOptions {
    /// Write a newly created revision owner in `conformance` instead of the
    /// existing owner's conformance (Transitional when the workbook has no
    /// revision owner).
    #[must_use]
    pub const fn conformance(mut self, conformance: super::model::RevisionConformance) -> Self {
        self.conformance = Some(conformance);
        self
    }
}

/// A staged, failure-atomic edit of the workbook revision owner of one
/// [`crate::Package`]. Dropping it discards it; [`Self::commit`] publishes a
/// new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    before: Snapshot,
    draft: Option<Revisions>,
    conformance: super::model::RevisionConformance,
}

impl Edit {
    /// Start an edit from the package's main workbook.
    fn start(target: crate::record::Source, options: EditOptions) -> Result<Self> {
        let before = Snapshot::load(&target)?;
        let draft = before.revisions().cloned();
        let conformance = options.conformance.unwrap_or_else(|| {
            before
                .conformance()
                .unwrap_or(super::model::RevisionConformance::Transitional)
        });
        Ok(Self {
            target,
            before,
            draft,
            conformance,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged revision package.
    #[must_use]
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
            draft.headers.properties.guid.clone_from(&last.guid);
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
            draft.headers.properties.guid.clone_from(&last.guid);
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

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.draft = self.before.revisions().cloned();
    }

    fn changed(&self) -> bool {
        self.before.revisions() != self.draft.as_ref()
    }

    /// Validate and publish the staged owner atomically.
    ///
    /// An edit without a change returns the source package, so a signed
    /// package keeps its signature byte for byte. A changed commit over a
    /// signed package is refused with [`crate::Error::Signed`]: remove the
    /// signature explicitly first ([`crate::Package::unsign`]).
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Signed`] and the publication's validation
    /// errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    pub(super) fn publish(self) -> Result<(Commit, Snapshot)> {
        if !self.changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                self.before,
            ));
        }
        crate::signature_policy::ensure_unsigned(&self.target)?;
        let mut candidate = OpcPackage::clone(&self.target);
        replace_workbook_revisions(&mut candidate, self.draft.as_ref(), self.conformance)?;
        let snapshot = Snapshot::load(&candidate)?;
        if !same_model(snapshot.revisions(), self.draft.as_ref()) {
            return Err(invalid(
                "revision package publication changed the staged model",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}

impl crate::Package {
    /// Start an edit of the workbook's revision metadata.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and the validation errors of the revision metadata graph.
    pub fn edit_revisions(&self) -> Result<Edit> {
        self.edit_revisions_with_options(EditOptions::default())
    }

    /// Start a revision edit with explicit options.
    ///
    /// # Errors
    ///
    /// See [`Self::edit_revisions`].
    pub fn edit_revisions_with_options(&self, options: EditOptions) -> Result<Edit> {
        Edit::start(crate::record::Source::new(self, "edit_revisions")?, options)
    }

    /// Apply an exact revision metadata patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_revisions(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_revisions")
    }
}

crate::record::raw_test_route!(edit_revisions);

/// Whether two revision owners are equal up to the order of XML attributes,
/// which the writer serializes canonically.
fn same_model(left: Option<&Revisions>, right: Option<&Revisions>) -> bool {
    match (left, right) {
        (None, None) => true,
        (Some(left), Some(right)) => left == right || canonical(left) == canonical(right),
        _ => false,
    }
}

fn canonical(value: &Revisions) -> Revisions {
    fn element(value: &mut super::model::RevisionXmlElement) {
        value.attributes.sort_by(|left, right| {
            (&left.namespace, &left.name).cmp(&(&right.namespace, &right.name))
        });
        value.children.iter_mut().for_each(element);
    }
    let mut value = value.clone();
    for log in &mut value.logs {
        for record in &mut log.log.records {
            record.attributes.sort_by(|left, right| {
                (&left.namespace, &left.name).cmp(&(&right.namespace, &right.name))
            });
            record.children.iter_mut().for_each(element);
        }
    }
    value
}
