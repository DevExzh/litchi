//! Immutable envelope snapshots and failure-atomic semantic edits.

use super::model::{Envelope, Message, Payload};
use crate::package::{Error as PackageError, Result as PackageResult};
use std::fmt;

/// An immutable optional `MsoEnvelopeCLSID` state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    envelope: Option<Envelope>,
}

impl Snapshot {
    /// Capture a present, validated envelope.
    pub fn new(envelope: Envelope) -> PackageResult<Self> {
        super::validation::validate(&envelope)?;
        Ok(Self {
            envelope: Some(envelope),
        })
    }

    /// Capture an absent FIB range.
    #[must_use]
    pub const fn empty() -> Self {
        Self { envelope: None }
    }

    pub(crate) fn from_option(envelope: Option<Envelope>) -> PackageResult<Self> {
        if let Some(value) = &envelope {
            super::validation::validate(value)?;
        }
        Ok(Self { envelope })
    }

    /// The typed envelope, when its FIB range is present.
    #[must_use]
    pub fn envelope(&self) -> Option<&Envelope> {
        self.envelope.as_ref()
    }

    /// The supported Office message body, when present under the known CLSID.
    #[must_use]
    pub fn message(&self) -> Option<&Message> {
        self.envelope.as_ref().and_then(Envelope::message)
    }

    /// Whether the FIB range is present.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.envelope.is_some()
    }

    /// Start an independent transaction from this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            working: self.clone(),
        }
    }
}

impl Envelope {
    /// Capture this envelope as a validated immutable snapshot.
    pub fn snapshot(&self) -> PackageResult<Snapshot> {
        Snapshot::new(self.clone())
    }

    /// Start a validated semantic transaction from this envelope.
    pub fn edit(&self) -> PackageResult<Transaction> {
        Ok(self.snapshot()?.edit())
    }
}

impl From<Envelope> for Snapshot {
    fn from(envelope: Envelope) -> Self {
        Self {
            envelope: Some(envelope),
        }
    }
}

/// A reversible semantic change between two envelope snapshots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// The exact semantic source snapshot required by this patch.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The exact semantic result snapshot produced by this patch.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the semantic state is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Return the exact inverse semantic patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone())
    }

    /// Apply this patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, TransactionError> {
        if source != &self.before {
            return Err(TransactionError::Conflict);
        }
        Ok(self.after.clone())
    }
}

/// The validated result of a semantic envelope transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    pub(crate) fn new(snapshot: Snapshot, patch: Patch) -> Self {
        Self { snapshot, patch }
    }

    /// The immutable post-edit snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split the result into its snapshot and reversible patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A mutable candidate built from an immutable envelope snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    before: Snapshot,
    working: Snapshot,
}

impl Transaction {
    /// Create a transaction from an immutable snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        Self {
            before: snapshot.clone(),
            working: snapshot,
        }
    }

    /// The current staged snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Whether the staged state differs from its source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before != self.working
    }

    /// Restore the staged state to the source snapshot.
    pub fn rollback(&mut self) {
        self.working = self.before.clone();
    }

    /// Replace the complete optional FIB range atomically.
    pub fn set(&mut self, envelope: Option<Envelope>) -> Result<(), TransactionError> {
        self.working = Snapshot::from_option(envelope).map_err(TransactionError::Invalid)?;
        Ok(())
    }

    /// Replace the present envelope, rejecting an absent range only at the
    /// package facade; detached transactions may also create the range.
    pub fn replace(&mut self, envelope: Envelope) -> Result<(), TransactionError> {
        self.set(Some(envelope))
    }

    /// Replace the supported Office message body under the known CLSID.
    pub fn set_message(&mut self, message: Message) -> Result<(), TransactionError> {
        self.replace(Envelope::from_message(message).map_err(TransactionError::Invalid)?)
    }

    /// Clone-first update of the supported Office message body.
    ///
    /// The closure never receives the transaction's live state. The edited
    /// clone is validated before it becomes the staged snapshot, so a failed
    /// edit leaves the transaction unchanged.
    pub fn update_message<F>(&mut self, edit: F) -> Result<(), TransactionError>
    where
        F: FnOnce(&mut Message),
    {
        let envelope = self
            .working
            .envelope
            .as_ref()
            .ok_or(TransactionError::Missing)?;
        let Payload::Message(message) = envelope.payload() else {
            return Err(TransactionError::Unsupported);
        };
        let mut message = message.as_ref().clone();
        edit(&mut message);
        self.set_message(message)
    }

    /// Remove the complete FIB range while retaining all other package bytes.
    pub fn clear(&mut self) {
        self.working = Snapshot::empty();
    }

    /// Commit the staged state into a reversible semantic patch.
    pub fn commit(self) -> Result<Commit, TransactionError> {
        let patch = Patch::new(self.before, self.working.clone());
        Ok(Commit::new(self.working, patch))
    }
}

/// Failure modes for semantic envelope edits.
#[derive(Debug)]
pub enum TransactionError {
    /// The candidate violates a bounded `[MS-DOC]`/`[MS-OSHARED]` invariant.
    Invalid(PackageError),
    /// The patch or package transaction was created from a different snapshot.
    Conflict,
    /// A message-only operation requires a present envelope.
    Missing,
    /// A message-only operation was requested for an unknown CLSID.
    Unsupported,
}

impl fmt::Display for TransactionError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("envelope transaction snapshot conflict"),
            Self::Missing => formatter.write_str("MsoEnvelope FIB range is absent"),
            Self::Unsupported => {
                formatter.write_str("the envelope CLSID does not select a supported message")
            },
        }
    }
}

impl std::error::Error for TransactionError {}

impl From<PackageError> for TransactionError {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}
