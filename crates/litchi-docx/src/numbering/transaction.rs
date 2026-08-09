#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Failure-atomic numbering snapshots, edits, and reversible patches.

use std::sync::Arc;

use crate::{Error, Result};

use super::codec;
use super::model::Collection;
use super::validation;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Change {
    pub(crate) abstract_num_id: u32,
    pub(crate) before: Option<bool>,
    pub(crate) after: Option<bool>,
}

/// An immutable, cheaply clonable source-preserving numbering snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    collection: Collection,
}

impl Snapshot {
    /// Parse and retain a bounded `numbering.xml` source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let collection = Collection::from_xml(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            collection,
        })
    }

    /// Borrow the authored numbering XML without copying it.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed numbering collection in source order.
    #[inline]
    #[must_use]
    pub const fn collection(&self) -> &Collection {
        &self.collection
    }

    /// Return one abstract numbering definition by its stable ID.
    #[inline]
    #[must_use]
    pub fn definition(&self, abstract_num_id: u32) -> Option<&super::Definition> {
        self.collection.get_abstract_num(abstract_num_id)
    }

    /// Return the optional section-break restart policy for one definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn restart_numbering_after_break(&self, abstract_num_id: u32) -> Result<Option<bool>> {
        self.definition(abstract_num_id)
            .map(super::model::Definition::restart_numbering_after_break)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "abstract numbering definition {abstract_num_id} does not exist"
                ))
            })
    }

    /// Start an isolated edit based on this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            changes: Vec::new(),
        }
    }
}

/// A numbering edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    changes: Vec<Change>,
}

impl Transaction {
    /// Return the projected section-break restart policy for one definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn restart_numbering_after_break(&self, abstract_num_id: u32) -> Result<Option<bool>> {
        let original = self.base.restart_numbering_after_break(abstract_num_id)?;
        Ok(self
            .changes
            .iter()
            .find(|change| change.abstract_num_id == abstract_num_id)
            .map_or(original, |change| change.after))
    }

    /// Set or remove the Word 2012 section-break restart attribute on one
    /// existing abstract numbering definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn set_restart_numbering_after_break(
        &mut self,
        abstract_num_id: u32,
        value: Option<bool>,
    ) -> Result<&mut Self> {
        validation::validate_restart_numbering_after_break(value)?;
        let before = self.base.restart_numbering_after_break(abstract_num_id)?;
        let current = self.restart_numbering_after_break(abstract_num_id)?;
        if current == value {
            return Ok(self);
        }

        if let Some(index) = self
            .changes
            .iter()
            .position(|change| change.abstract_num_id == abstract_num_id)
        {
            if before == value {
                self.changes.remove(index);
            } else {
                self.changes[index].after = value;
            }
        } else {
            self.changes.push(Change {
                abstract_num_id,
                before,
                after: value,
            });
        }
        Ok(self)
    }

    /// Remove the direct section-break restart attribute from one definition.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn clear_restart_numbering_after_break(
        &mut self,
        abstract_num_id: u32,
    ) -> Result<&mut Self> {
        self.set_restart_numbering_after_break(abstract_num_id, None)
    }

    /// Validate and publish the edit without changing the source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn commit(self) -> Result<Commit> {
        if self.changes.is_empty() {
            let patch = Patch {
                before: Vec::new(),
                after: Vec::new(),
                before_xml: self.base.xml.clone(),
                after_xml: self.base.xml.clone(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
            });
        }

        let xml =
            codec::rewrite_restart_numbering_after_break(self.base.xml_bytes(), &self.changes)?;
        let snapshot = Snapshot::from_xml(xml)?;
        let patch = Patch {
            before: self.changes.clone(),
            after: self
                .changes
                .iter()
                .map(|change| Change {
                    abstract_num_id: change.abstract_num_id,
                    before: change.before,
                    after: change.after,
                })
                .collect(),
            before_xml: self.base.xml,
            after_xml: snapshot.xml.clone(),
        };
        Ok(Commit { snapshot, patch })
    }
}

/// A successful publication containing the new snapshot and reversible patch.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A source-checked, reversible abstract-definition restart patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Vec<Change>,
    after: Vec<Change>,
    before_xml: Arc<[u8]>,
    after_xml: Arc<[u8]>,
}

impl Patch {
    /// Return the expected policy before applying the patch.
    #[must_use]
    pub fn before_restart_numbering_after_break(
        &self,
        abstract_num_id: u32,
    ) -> Option<Option<bool>> {
        self.before
            .iter()
            .find(|change| change.abstract_num_id == abstract_num_id)
            .map(|change| change.before)
    }

    /// Return the policy produced by the patch.
    #[must_use]
    pub fn after_restart_numbering_after_break(
        &self,
        abstract_num_id: u32,
    ) -> Option<Option<bool>> {
        self.after
            .iter()
            .find(|change| change.abstract_num_id == abstract_num_id)
            .map(|change| change.after)
    }

    /// Return the inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self
                .after
                .iter()
                .map(|change| Change {
                    abstract_num_id: change.abstract_num_id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            after: self
                .before
                .iter()
                .map(|change| Change {
                    abstract_num_id: change.abstract_num_id,
                    before: change.after,
                    after: change.before,
                })
                .collect(),
            before_xml: self.after_xml.clone(),
            after_xml: self.before_xml.clone(),
        }
    }

    /// Apply the patch only when the target has the exact source bytes and
    /// semantic values captured when it was created.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_ref() != self.before_xml.as_ref() {
            return Err(Error::InvalidFormat(
                "numbering patch source snapshot does not match its byte precondition".into(),
            ));
        }
        for change in &self.before {
            let current = source
                .definition(change.abstract_num_id)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "numbering patch references missing abstract definition {}",
                        change.abstract_num_id
                    ))
                })?
                .restart_numbering_after_break();
            if current != change.before {
                return Err(Error::InvalidFormat(
                    "numbering patch source snapshot does not match its semantic precondition"
                        .into(),
                ));
            }
        }
        if self.before.is_empty() {
            return Ok(source.clone());
        }
        Snapshot::from_xml(self.after_xml.to_vec())
    }
}
