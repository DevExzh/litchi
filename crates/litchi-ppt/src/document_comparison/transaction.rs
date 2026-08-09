//! Immutable document-comparison snapshots and atomic review edits.

use super::codec;
use super::model::{DiffFlags, Entry, Limits, Review, ReviewingToolbarStates, Unknown};
use super::validation;
use crate::package::{Error, Result};
use crate::records::Record;
use std::sync::Arc;

/// A deterministic content revision for optimistic review integration.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }

    pub(crate) fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }
}

/// The review metadata facet changed by one transaction operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChangeKind {
    Toolbar,
    ReviewerName,
    DocumentFlags,
}

/// One semantic change to the inert review metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Change {
    Toolbar {
        before: Option<ReviewingToolbarStates>,
        after: Option<ReviewingToolbarStates>,
    },
    ReviewerName {
        tree_index: usize,
        before: String,
        after: String,
    },
    DocumentFlags {
        tree_index: usize,
        before: DiffFlags,
        after: DiffFlags,
    },
}

impl Change {
    #[must_use]
    pub const fn kind(&self) -> ChangeKind {
        match self {
            Self::Toolbar { .. } => ChangeKind::Toolbar,
            Self::ReviewerName { .. } => ChangeKind::ReviewerName,
            Self::DocumentFlags { .. } => ChangeKind::DocumentFlags,
        }
    }

    #[must_use]
    pub const fn toolbar(
        &self,
    ) -> Option<(
        Option<ReviewingToolbarStates>,
        Option<ReviewingToolbarStates>,
    )> {
        match self {
            Self::Toolbar { before, after } => Some((*before, *after)),
            Self::ReviewerName { .. } | Self::DocumentFlags { .. } => None,
        }
    }

    #[must_use]
    pub const fn tree_index(&self) -> Option<usize> {
        match self {
            Self::ReviewerName { tree_index, .. } | Self::DocumentFlags { tree_index, .. } => {
                Some(*tree_index)
            },
            Self::Toolbar { .. } => None,
        }
    }
}

/// A reversible revision-checked patch produced by a successful edit.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    changes: Vec<Change>,
    limits: Limits,
}

impl Patch {
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
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    /// Undo this patch against its exact committed target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.target || current.bytes() != self.after.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot undo review edits against a different revision".into(),
            ));
        }
        Snapshot::parse_with_limits(&self.before, self.limits)
    }

    /// Redo this patch against its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.base || current.bytes() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "cannot redo review edits against a different revision".into(),
            ));
        }
        Snapshot::parse_with_limits(&self.after, self.limits)
    }
}

/// An immutable, lossless snapshot of one `DocumentContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) root: Record,
    bytes: Arc<[u8]>,
    limits: Limits,
}

impl Snapshot {
    /// Capture a validated document record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record(root: Record) -> Result<Self> {
        Self::from_record_with_limits(root, Limits::default())
    }

    /// Capture a document record with explicit bounded-resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "taking the record by value is the established public API; the snapshot stores a normalized re-parse, so the argument is only borrowed internally"
    )]
    pub fn from_record_with_limits(root: Record, limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        validation::validate_document(&root, validated)?;
        let bytes = codec::encode_document(&root)?;
        if bytes.len() > validated.max_bytes {
            return Err(Error::InvalidFormat(
                "document-comparison snapshot exceeds the byte limit".into(),
            ));
        }
        let (parsed, consumed) = Record::parse_strict(&bytes, 0)?;
        if consumed != bytes.len() {
            return Err(Error::Corrupted(
                "document-comparison snapshot has trailing bytes".into(),
            ));
        }
        validation::validate_document(&parsed, validated)?;
        let _ = codec::read_review(&parsed, validated)?;
        Ok(Self {
            root: parsed,
            bytes: Arc::from(bytes.into_boxed_slice()),
            limits: validated,
        })
    }

    /// Parse exactly one complete `DocumentContainer` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one document under explicit bounded-resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let snapshot = Self::from_record_with_limits(
            {
                let (record, consumed) = Record::parse_strict(bytes, 0)?;
                if consumed != bytes.len() {
                    return Err(Error::Corrupted(
                        "document-comparison input contains trailing bytes".into(),
                    ));
                }
                record
            },
            limits,
        )?;
        if snapshot.bytes() != bytes {
            return Err(Error::Corrupted(
                "document-comparison snapshot is not losslessly representable".into(),
            ));
        }
        Ok(snapshot)
    }

    #[must_use]
    pub const fn record(&self) -> &Record {
        &self.root
    }
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the typed review records and opaque records in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn review(&self) -> Result<Review> {
        codec::read_review(&self.root, self.limits)
    }

    /// Return the typed reviewing-toolbar state, when present.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        Ok(self.review()?.toolbar())
    }

    /// Return the `index`th reviewer tree in the document-comparison payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn diff_tree(&self, index: usize) -> Result<Option<super::model::DiffTree10>> {
        Ok(self.review()?.diff_tree(index).cloned())
    }

    /// Return opaque records retained by the review owner.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn unknown_records(&self) -> Result<Vec<Unknown>> {
        Ok(self.review()?.unknown_records().cloned().collect())
    }

    #[must_use]
    pub fn revision(&self) -> Revision {
        Revision::from_bytes(&self.bytes)
    }

    /// Start an isolated semantic edit.
    #[must_use]
    pub fn edit(&self) -> Editor {
        Editor {
            source: self.clone(),
            root: self.root.clone(),
            changes: Vec::new(),
        }
    }
}

/// An isolated atomic edit over one document's review metadata.
#[derive(Debug, Clone)]
pub struct Editor {
    source: Snapshot,
    root: Record,
    changes: Vec<Change>,
}

impl Editor {
    #[must_use]
    pub const fn record(&self) -> &Record {
        &self.root
    }
    /// Return the typed review records and opaque records in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the inert review payload is missing or malformed.
    pub fn review(&self) -> Result<Review> {
        codec::read_review(&self.root, self.source.limits)
    }
    /// Return the typed reviewing-toolbar state, when present.
    ///
    /// # Errors
    ///
    /// Returns an error if the inert review payload is missing or malformed.
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        Ok(self.review()?.toolbar())
    }
    /// Return the `index`th reviewer tree in the document-comparison payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the inert review payload is missing or malformed.
    pub fn diff_tree(&self, index: usize) -> Result<Option<super::model::DiffTree10>> {
        Ok(self.review()?.diff_tree(index).cloned())
    }
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        !self.changes.is_empty()
    }

    /// Set or replace the reviewing toolbar atom without touching other records.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_toolbar(&mut self, value: ReviewingToolbarStates) -> Result<()> {
        self.replace_toolbar(Some(value))
    }

    /// Replace one reviewer name in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_reviewer_name(&mut self, tree_index: usize, value: impl Into<String>) -> Result<()> {
        let name = value.into();
        validation::validate_reviewer_name(&name)?;
        let mut review = self.review()?;
        let entry = review
            .entries
            .iter_mut()
            .filter(|entry| matches!(entry, Entry::Diff(_)))
            .nth(tree_index)
            .ok_or_else(|| Error::InvalidFormat("reviewer tree index is out of range".into()))?;
        let Entry::Diff(tree) = entry else {
            unreachable!("filtered reviewer tree entry")
        };
        if tree.reviewer_name == name {
            return Ok(());
        }
        let before = std::mem::replace(&mut tree.reviewer_name, name.clone());
        codec::write_review(&mut self.root, &review, self.source.limits)?;
        self.changes.push(Change::ReviewerName {
            tree_index,
            before,
            after: name,
        });
        Ok(())
    }

    /// Replace the document-level display flags of one reviewer tree.
    ///
    /// This changes only what the legacy reviewing UI displays. It never
    /// applies, rejects, or generates any underlying presentation change.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_document_flags(&mut self, tree_index: usize, value: DiffFlags) -> Result<()> {
        let mut review = self.review()?;
        let entry = review
            .entries
            .iter_mut()
            .filter(|entry| matches!(entry, Entry::Diff(_)))
            .nth(tree_index)
            .ok_or_else(|| Error::InvalidFormat("reviewer tree index is out of range".into()))?;
        let Entry::Diff(tree) = entry else {
            unreachable!("filtered reviewer tree entry")
        };
        let before = tree.document_flags();
        if before == value {
            return Ok(());
        }
        tree.document_diff.set_flags(value)?;
        codec::write_review(&mut self.root, &review, self.source.limits)?;
        self.changes.push(Change::DocumentFlags {
            tree_index,
            before,
            after: value,
        });
        Ok(())
    }

    /// Remove the reviewing toolbar atom, preserving the rest of the payload.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn clear_toolbar(&mut self) -> Result<bool> {
        let before = self.toolbar()?;
        if before.is_none() {
            return Ok(false);
        }
        self.replace_toolbar(None)?;
        Ok(true)
    }

    fn replace_toolbar(&mut self, after: Option<ReviewingToolbarStates>) -> Result<()> {
        let before = self.toolbar()?;
        if before == after {
            return Ok(());
        }
        let mut review = self.review()?;
        if let Some(index) = review
            .entries
            .iter()
            .position(|entry| matches!(entry, Entry::Toolbar(_)))
        {
            if let Some(value) = after {
                review.entries[index] = Entry::Toolbar(value);
            } else {
                review.entries.remove(index);
            }
        } else if let Some(value) = after {
            let index = review
                .entries
                .iter()
                .position(|entry| !matches!(entry, Entry::Unknown(_)))
                .unwrap_or(review.entries.len());
            review.entries.insert(index, Entry::Toolbar(value));
        }
        codec::write_review(&mut self.root, &review, self.source.limits)?;
        self.changes.push(Change::Toolbar { before, after });
        Ok(())
    }

    /// Capture the candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::from_record_with_limits(self.root.clone(), self.source.limits)
    }

    /// Validate and publish the candidate and its reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = Snapshot::from_record_with_limits(self.root, self.source.limits)?;
        let patch = Patch {
            base: self.source.revision(),
            target: snapshot.revision(),
            before: self.source.bytes.clone(),
            after: snapshot.bytes.clone(),
            changes: self.changes,
            limits: self.source.limits,
        };
        Ok(Commit { snapshot, patch })
    }

    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }
}

/// A successful immutable target and reversible review patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    /// Return the typed reviewing-toolbar state of the committed snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the inert review payload is missing or malformed.
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        self.snapshot.toolbar()
    }
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
