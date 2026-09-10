//! Package-facing owner for a Word FIB and its selected table stream.

use super::FileInformationBlock;
use super::transaction::{Commit, Snapshot, Transaction};
use crate::package::{Error as PackageError, Result};
use crate::parts::protection::ProtectionPolicy;
use litchi_codepage::Ansi;
use litchi_ole_common::smart_tags::Limits;

/// A small package facade that publishes only fixed-topology table edits.
///
/// The containing OLE2 storage remains outside this owner. Callers can put
/// [`Self::finish`] back into the existing `WordDocument`/table streams while
/// retaining every unrelated compound-file stream and FIB pointer.
#[derive(Debug, Clone)]
pub struct Editor {
    current: Snapshot,
}

impl Editor {
    /// Open the smart-tag view from an already selected Word table stream.
    pub fn open(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        Ok(Self {
            current: Snapshot::parse(fib, table_stream)?,
        })
    }

    /// Open with an explicit policy for changed protected metadata.
    pub fn open_with_policy(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        Ok(Self {
            current: Snapshot::parse_with_policy(fib, table_stream, protection_policy)?,
        })
    }

    /// Open with an explicit ANSI page and default smart-tag limits.
    pub fn open_with_ansi(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        ansi: Ansi,
    ) -> Result<Self> {
        Ok(Self {
            current: Snapshot::parse_with(fib, table_stream, ansi)?,
        })
    }

    /// Open with custom bounded parsing and LCID-derived ANSI decoding.
    pub fn open_with_limits(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        limits: Limits,
    ) -> Result<Self> {
        Ok(Self {
            current: Snapshot::parse_with_limits(fib, table_stream, limits)?,
        })
    }

    /// Open with bounded parsing and an explicit protected-edit policy.
    pub fn open_with_limits_and_policy(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        limits: Limits,
        protection_policy: ProtectionPolicy,
    ) -> Result<Self> {
        Ok(Self {
            current: Snapshot::parse_with_options_and_policy(
                fib,
                table_stream,
                None,
                limits,
                protection_policy,
            )?,
        })
    }

    /// The current immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.current
    }

    /// Start an isolated transaction from the current package state.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.current.edit()
    }

    /// Apply and publish a transaction only if it was created from this state.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction
            .commit()
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let current = commit
            .patch()
            .apply(&self.current)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        self.current = current;
        Ok(commit)
    }

    /// Return the FIB and selected table stream for package publication.
    #[must_use]
    pub fn finish(&self) -> (Vec<u8>, Vec<u8>) {
        self.current.finish()
    }
}
