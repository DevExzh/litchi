//! Immutable document-comparison snapshots and atomic review edits.

use super::codec;
use super::model::{Entry, Limits, Review, ReviewingToolbarStates, Unknown};
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

/// A semantic change to the review-owned toolbar atom.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change {
    before: Option<ReviewingToolbarStates>,
    after: Option<ReviewingToolbarStates>,
}

impl Change {
    pub const fn before(&self) -> Option<ReviewingToolbarStates> {
        self.before
    }
    pub const fn after(&self) -> Option<ReviewingToolbarStates> {
        self.after
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
    pub const fn base(&self) -> Revision {
        self.base
    }
    pub const fn target(&self) -> Revision {
        self.target
    }
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }
    pub fn before_bytes(&self) -> &[u8] {
        &self.before
    }
    pub fn after_bytes(&self) -> &[u8] {
        &self.after
    }

    /// Undo this patch against its exact committed target snapshot.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.target {
            return Err(Error::InvalidFormat(
                "cannot undo review edits against a different revision".into(),
            ));
        }
        Snapshot::parse_with_limits(&self.before, self.limits)
    }

    /// Redo this patch against its exact source snapshot.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision() != self.base {
            return Err(Error::InvalidFormat(
                "cannot redo review edits against a different revision".into(),
            ));
        }
        Snapshot::parse_with_limits(&self.after, self.limits)
    }
}

/// An immutable, lossless snapshot of one DocumentContainer.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) root: Record,
    bytes: Arc<[u8]>,
    limits: Limits,
}

impl Snapshot {
    /// Capture a validated document record.
    pub fn from_record(root: Record) -> Result<Self> {
        Self::from_record_with_limits(root, Limits::default())
    }

    /// Capture a document record with explicit bounded-resource limits.
    pub fn from_record_with_limits(root: Record, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        validation::validate_document(&root, limits)?;
        let bytes = codec::encode_document(&root)?;
        if bytes.len() > limits.max_bytes {
            return Err(Error::InvalidFormat(
                "document-comparison snapshot exceeds the byte limit".into(),
            ));
        }
        let (parsed, consumed) = Record::parse(&bytes, 0)?;
        if consumed != bytes.len() {
            return Err(Error::Corrupted(
                "document-comparison snapshot has trailing bytes".into(),
            ));
        }
        validation::validate_document(&parsed, limits)?;
        let _ = codec::read_review(&parsed, limits)?;
        Ok(Self {
            root: parsed,
            bytes: Arc::from(bytes.into_boxed_slice()),
            limits,
        })
    }

    /// Parse exactly one complete DocumentContainer record.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one document under explicit bounded-resource limits.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let snapshot = Self::from_record_with_limits(
            {
                let (record, consumed) = Record::parse(bytes, 0)?;
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

    pub const fn record(&self) -> &Record {
        &self.root
    }
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return the typed review records and opaque records in source order.
    pub fn review(&self) -> Result<Review> {
        codec::read_review(&self.root, self.limits)
    }

    /// Return the typed reviewing-toolbar state, when present.
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        Ok(self.review()?.toolbar())
    }

    /// Return opaque records retained by the review owner.
    pub fn unknown_records(&self) -> Result<Vec<Unknown>> {
        Ok(self.review()?.unknown_records().cloned().collect())
    }

    #[must_use]
    pub fn revision(&self) -> Revision {
        Revision::from_bytes(&self.bytes)
    }

    /// Start an isolated semantic edit.
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
    pub const fn record(&self) -> &Record {
        &self.root
    }
    pub fn review(&self) -> Result<Review> {
        codec::read_review(&self.root, self.source.limits)
    }
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        Ok(self.review()?.toolbar())
    }
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }
    pub const fn is_changed(&self) -> bool {
        !self.changes.is_empty()
    }

    /// Set or replace the reviewing toolbar atom without touching other records.
    pub fn set_toolbar(&mut self, value: ReviewingToolbarStates) -> Result<()> {
        self.replace_toolbar(Some(value))
    }

    /// Remove the reviewing toolbar atom, preserving the rest of the payload.
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
        self.changes.push(Change { before, after });
        Ok(())
    }

    /// Capture the candidate without publishing it.
    pub fn snapshot(&self) -> Result<Snapshot> {
        Snapshot::from_record_with_limits(self.root.clone(), self.source.limits)
    }

    /// Validate and publish the candidate and its reversible patch.
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
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    pub fn toolbar(&self) -> Result<Option<ReviewingToolbarStates>> {
        self.snapshot.toolbar()
    }
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}
