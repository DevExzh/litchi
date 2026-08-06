//! Immutable route-slip snapshots and semantic transactions.

use super::model::{Metadata, Protection, Recipient};
use super::validation;
use crate::package::{Error as PackageError, Result as PackageResult};
use std::fmt;

/// A recipient target resolved against one immutable metadata snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecipientSelector<'a> {
    /// The recipient currently named by `Metadata::stage`.
    Current,
    /// A zero-based recipient index.
    Index(usize),
    /// An exact, case-sensitive ANSI byte name.
    Name(&'a [u8]),
}

/// A checked recipient-selection failure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum RecipientSelectionError {
    /// The snapshot has no route-slip metadata.
    NoMetadata,
    /// The requested index is outside the recipient list.
    IndexOutOfBounds { index: usize, len: usize },
    /// No recipient has the exact requested ANSI name.
    NameNotFound { name: Vec<u8> },
    /// More than one recipient has the exact requested ANSI name.
    AmbiguousName { name: Vec<u8>, matches: Vec<usize> },
}

impl fmt::Display for RecipientSelectionError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NoMetadata => f.write_str("route-slip metadata is missing"),
            Self::IndexOutOfBounds { index, len } => {
                write!(f, "recipient index {index} is outside {len} recipients")
            },
            Self::NameNotFound { name } => write!(f, "recipient name {name:?} was not found"),
            Self::AmbiguousName { name, matches } => {
                write!(f, "recipient name {name:?} is ambiguous at {matches:?}")
            },
        }
    }
}

impl std::error::Error for RecipientSelectionError {}

/// Errors produced while staging semantic route-slip edits.
#[derive(Debug)]
pub enum Error {
    /// A recipient selector could not be resolved.
    Selection(RecipientSelectionError),
    /// The route-slip protection policy does not permit metadata editing.
    Protected(Protection),
    /// The candidate metadata violates the strict MS-DOC codec constraints.
    Invalid(PackageError),
    /// The transaction was created from a different snapshot.
    Conflict,
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Selection(error) => error.fmt(f),
            Self::Protected(protection) => {
                write!(f, "route-slip metadata is protected by {protection:?}")
            },
            Self::Invalid(error) => error.fmt(f),
            Self::Conflict => f.write_str("route-slip transaction snapshot conflict"),
        }
    }
}

impl std::error::Error for Error {}

impl From<RecipientSelectionError> for Error {
    fn from(error: RecipientSelectionError) -> Self {
        Self::Selection(error)
    }
}

impl From<PackageError> for Error {
    fn from(error: PackageError) -> Self {
        Self::Invalid(error)
    }
}

/// An immutable semantic route-slip state. `None` means no route-slip range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    metadata: Option<Metadata>,
}

impl Snapshot {
    /// Creates a snapshot containing route-slip metadata.
    pub fn new(metadata: Metadata) -> PackageResult<Self> {
        validation::metadata(&metadata)?;
        Ok(Self {
            metadata: Some(metadata),
        })
    }

    /// Creates an empty snapshot with no route-slip metadata.
    #[must_use]
    pub const fn empty() -> Self {
        Self { metadata: None }
    }

    pub(crate) fn from_option(metadata: Option<Metadata>) -> PackageResult<Self> {
        if let Some(value) = &metadata {
            validation::metadata(value)?;
        }
        Ok(Self { metadata })
    }

    /// Returns the immutable route-slip metadata, when present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.metadata.as_ref()
    }

    /// Whether this snapshot contains route-slip metadata.
    #[must_use]
    pub fn is_present(&self) -> bool {
        self.metadata.is_some()
    }

    /// Starts an independent transaction from this snapshot.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            before: self.clone(),
            working: self.clone(),
        }
    }

    /// Resolves a checked recipient selector against this immutable snapshot.
    pub fn resolve(
        &self,
        selector: RecipientSelector<'_>,
    ) -> Result<usize, RecipientSelectionError> {
        let metadata = self
            .metadata
            .as_ref()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        match selector {
            RecipientSelector::Current => Ok(usize::from(metadata.stage)),
            RecipientSelector::Index(index) => {
                if index < metadata.recipients.len() {
                    Ok(index)
                } else {
                    Err(RecipientSelectionError::IndexOutOfBounds {
                        index,
                        len: metadata.recipients.len(),
                    })
                }
            },
            RecipientSelector::Name(name) => {
                let matches = metadata
                    .recipients
                    .iter()
                    .enumerate()
                    .filter_map(|(index, recipient)| {
                        (recipient.name.as_bytes() == name).then_some(index)
                    })
                    .collect::<Vec<_>>();
                match matches.as_slice() {
                    [] => Err(RecipientSelectionError::NameNotFound {
                        name: name.to_vec(),
                    }),
                    [index] => Ok(*index),
                    _ => Err(RecipientSelectionError::AmbiguousName {
                        name: name.to_vec(),
                        matches,
                    }),
                }
            },
        }
    }
}

impl Metadata {
    /// Creates a validated immutable snapshot of this metadata.
    pub fn snapshot(&self) -> PackageResult<Snapshot> {
        Snapshot::new(self.clone())
    }

    /// Starts a validated semantic transaction from this metadata.
    pub fn edit(&self) -> PackageResult<Transaction> {
        Ok(self.snapshot()?.edit())
    }
}

impl From<Metadata> for Snapshot {
    fn from(metadata: Metadata) -> Self {
        Self {
            metadata: Some(metadata),
        }
    }
}

/// A reversible semantic change between two route-slip snapshots.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    pub(crate) fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// The exact semantic source snapshot.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// The exact semantic result snapshot.
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether the semantic state did not change.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Returns the exact inverse semantic patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self::new(self.after.clone(), self.before.clone())
    }

    /// Applies this patch only to its expected source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source != &self.before {
            return Err(Error::Conflict);
        }
        Ok(self.after.clone())
    }
}

/// The validated result of a semantic route-slip transaction.
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

    /// Splits the commit into its snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A mutable candidate built from an immutable route-slip snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    before: Snapshot,
    working: Snapshot,
}

impl Transaction {
    /// Creates a transaction from an immutable snapshot.
    #[must_use]
    pub fn new(snapshot: Snapshot) -> Self {
        Self {
            before: snapshot.clone(),
            working: snapshot,
        }
    }

    /// Returns the current candidate snapshot without publishing it.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.working
    }

    /// Restores the transaction candidate to its source snapshot.
    pub fn rollback(&mut self) {
        self.working = self.before.clone();
    }

    /// Whether the candidate differs from its source snapshot.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before != self.working
    }

    /// Replaces the route-slip metadata candidate.
    pub fn set(&mut self, metadata: Metadata) -> Result<(), Error> {
        self.ensure_editable()?;
        validation::metadata(&metadata)?;
        self.working = Snapshot::from(metadata);
        Ok(())
    }

    /// Sets the current stage by a checked recipient selector.
    pub fn set_stage(&mut self, selector: RecipientSelector<'_>) -> Result<(), Error> {
        let index = self.resolve(selector)?;
        self.ensure_editable()?;
        let metadata = self
            .working
            .metadata
            .as_mut()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        metadata.stage = u16::try_from(index).map_err(|_| {
            Error::Invalid(PackageError::Corrupted(
                "recipient index exceeds u16::MAX".into(),
            ))
        })?;
        Ok(())
    }

    /// Advances to the next recipient without leaving the checked stage
    /// domain. The final recipient cannot advance further.
    pub fn advance_stage(&mut self) -> Result<usize, Error> {
        self.ensure_editable()?;
        let metadata = self
            .working
            .metadata
            .as_mut()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        let next = usize::from(metadata.stage).checked_add(1).ok_or_else(|| {
            Error::Invalid(PackageError::Corrupted("route stage overflows".into()))
        })?;
        if next >= metadata.recipients.len() {
            return Err(Error::Invalid(PackageError::Corrupted(
                "route slip is already at its final recipient".into(),
            )));
        }
        metadata.stage = u16::try_from(next).map_err(|_| {
            Error::Invalid(PackageError::Corrupted(
                "recipient index exceeds u16::MAX".into(),
            ))
        })?;
        Ok(next)
    }

    /// Adds a recipient after the existing ordered recipients.
    pub fn add_recipient(&mut self, recipient: Recipient) -> Result<(), Error> {
        self.ensure_editable()?;
        validation::recipient(&recipient)?;
        let metadata = self
            .working
            .metadata
            .as_mut()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        metadata.recipients.push(recipient);
        validation::metadata(metadata)?;
        Ok(())
    }

    /// Replaces one recipient while retaining its position.
    pub fn replace_recipient(
        &mut self,
        selector: RecipientSelector<'_>,
        recipient: Recipient,
    ) -> Result<(), Error> {
        let index = self.resolve(selector)?;
        self.ensure_editable()?;
        validation::recipient(&recipient)?;
        let metadata = self
            .working
            .metadata
            .as_mut()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        metadata.recipients[index] = recipient;
        validation::metadata(metadata)?;
        Ok(())
    }

    /// Removes one recipient and keeps the current stage valid.
    pub fn remove_recipient(&mut self, selector: RecipientSelector<'_>) -> Result<(), Error> {
        let index = self.resolve(selector)?;
        self.ensure_editable()?;
        let metadata = self
            .working
            .metadata
            .as_mut()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        if metadata.recipients.len() == 1 {
            return Err(Error::Invalid(PackageError::Corrupted(
                "a route slip must retain at least one recipient".into(),
            )));
        }
        metadata.recipients.remove(index);
        if usize::from(metadata.stage) == index {
            metadata.stage = metadata
                .stage
                .min(u16::try_from(metadata.recipients.len() - 1).unwrap_or(u16::MAX));
        } else if usize::from(metadata.stage) > index {
            metadata.stage -= 1;
        }
        validation::metadata(metadata)?;
        Ok(())
    }

    /// Clears the route-slip metadata from the candidate package state.
    pub fn clear(&mut self) -> Result<(), Error> {
        self.ensure_editable()?;
        self.working = Snapshot::empty();
        Ok(())
    }

    /// Completes routing and removes the route-slip metadata from the candidate.
    pub fn complete(&mut self) -> Result<(), Error> {
        self.clear()
    }

    /// Atomically publishes the candidate as a snapshot and reversible patch.
    pub fn commit(self) -> Result<Commit, Error> {
        validation::metadata_option(&self.working.metadata)?;
        let patch = Patch::new(self.before, self.working.clone());
        Ok(Commit::new(self.working, patch))
    }

    fn resolve(&self, selector: RecipientSelector<'_>) -> Result<usize, RecipientSelectionError> {
        let metadata = self
            .working
            .metadata
            .as_ref()
            .ok_or(RecipientSelectionError::NoMetadata)?;
        match selector {
            RecipientSelector::Current => Ok(usize::from(metadata.stage)),
            RecipientSelector::Index(index) => {
                if index < metadata.recipients.len() {
                    Ok(index)
                } else {
                    Err(RecipientSelectionError::IndexOutOfBounds {
                        index,
                        len: metadata.recipients.len(),
                    })
                }
            },
            RecipientSelector::Name(name) => {
                let matches = metadata
                    .recipients
                    .iter()
                    .enumerate()
                    .filter_map(|(index, recipient)| {
                        (recipient.name.as_bytes() == name).then_some(index)
                    })
                    .collect::<Vec<_>>();
                match matches.as_slice() {
                    [] => Err(RecipientSelectionError::NameNotFound {
                        name: name.to_vec(),
                    }),
                    [index] => Ok(*index),
                    _ => Err(RecipientSelectionError::AmbiguousName {
                        name: name.to_vec(),
                        matches,
                    }),
                }
            },
        }
    }

    fn ensure_editable(&self) -> Result<(), Error> {
        if let Some(metadata) = self.working.metadata() {
            validation::editable(metadata).map_err(Error::Protected)
        } else {
            Ok(())
        }
    }
}
