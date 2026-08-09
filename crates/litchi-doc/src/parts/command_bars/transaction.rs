//! Immutable command-bar snapshots and failure-atomic semantic edits.

use super::model::CommandBars;
use crate::package::{Error as PackageError, Result as PackageResult};

/// An immutable owned snapshot of the optional `fcCmds` command-bar range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    command_bars: Option<CommandBars<'static>>,
}

impl Snapshot {
    /// Creates a validated owned snapshot from a decoded command-bar range.
    pub fn new(command_bars: Option<CommandBars<'_>>) -> PackageResult<Self> {
        let command_bars = command_bars.map(CommandBars::into_owned);
        if let Some(value) = &command_bars {
            value.validate()?;
        }
        Ok(Self { command_bars })
    }

    /// Creates an empty snapshot with `fcCmds/lcbCmds` absent.
    #[must_use]
    pub const fn empty() -> Self {
        Self { command_bars: None }
    }

    /// Returns the owned command-bar metadata, when the FIB range is present.
    #[must_use]
    pub fn command_bars(&self) -> Option<&CommandBars<'static>> {
        self.command_bars.as_ref()
    }

    /// Whether the command-bar FIB range is present.
    #[must_use]
    pub const fn is_present(&self) -> bool {
        self.command_bars.is_some()
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

impl CommandBars<'_> {
    /// Captures this decoded command-bar range as an owned snapshot.
    pub fn snapshot(&self) -> PackageResult<Snapshot> {
        Snapshot::new(Some(self.clone()))
    }

    /// Starts a validated semantic transaction from this command-bar range.
    pub fn edit(&self) -> PackageResult<Transaction> {
        Ok(self.snapshot()?.edit())
    }
}

/// A staged, reversible command-bar metadata edit.
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

    /// Replaces the complete optional command-bar range atomically.
    pub fn set(&mut self, command_bars: Option<CommandBars<'_>>) -> Result<(), Error> {
        self.working = Snapshot::new(command_bars).map_err(Error::Invalid)?;
        Ok(())
    }

    /// Replaces the complete command-bar range while keeping it present.
    pub fn replace(&mut self, command_bars: CommandBars<'_>) -> Result<(), Error> {
        self.set(Some(command_bars))
    }

    /// Removes the command-bar FIB range from the staged snapshot.
    pub fn clear(&mut self) {
        self.working = Snapshot::empty();
    }

    /// Commits the staged edit and returns its reversible semantic patch.
    pub fn commit(self) -> Result<Commit, Error> {
        Ok(Commit {
            before: self.before,
            after: self.working,
        })
    }
}

/// A reversible change between two command-bar snapshots.
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

    /// Applies this patch only to the exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot, Error> {
        if source != &self.before {
            return Err(Error::Conflict);
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

/// Errors produced while staging or applying a command-bar edit.
#[derive(Debug)]
pub enum Error {
    /// The candidate violates a bounded command-bar invariant.
    Invalid(PackageError),
    /// The transaction was created from a different snapshot.
    Conflict,
}

impl std::fmt::Display for Error {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::Conflict => formatter.write_str("command-bar transaction snapshot conflict"),
        }
    }
}

impl std::error::Error for Error {}
