//! FIB/table-stream package integration for command-bar records.

use super::codec;
use super::model::CommandBars;
use super::transaction::{
    Commit as TransactionCommit, Error as TransactionError, Snapshot as TransactionSnapshot,
    Transaction,
};
use super::validation;
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use crate::writer::fib::FibBuilder;
use litchi_ole_common::object::{Editor as ObjectEditor, Limits, Patch as ObjectPatch, Targets};

/// `FibRgFcLcb97` index of fcCmds/lcbCmds (MS-DOC 2.5).
pub const FIB_INDEX_CMDS: usize = 24;

/// Parse the optional command-bar table addressed by fcCmds/lcbCmds.
pub fn parse<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
) -> Result<Option<CommandBars<'a>>> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_CMDS) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }

    let start = usize::try_from(offset).map_err(|_| corrupted("fcCmds exceeds usize"))?;
    let length = usize::try_from(length).map_err(|_| corrupted("lcbCmds exceeds usize"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("fcCmds/lcbCmds range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("Tcg extends beyond the table stream"))?;
    codec::parse_bytes(data).map(Some)
}

/// Append a command-bar table and update its FIB pointer atomically with the
/// in-memory table-stream operation.
pub fn write(
    fib: &mut FibBuilder,
    table_stream: &mut Vec<u8>,
    value: &CommandBars<'_>,
) -> Result<()> {
    let data = codec::to_bytes(value)?;
    let offset = u32::try_from(table_stream.len())
        .map_err(|_| corrupted("table-stream offset exceeds u32::MAX"))?;
    let length = u32::try_from(data.len()).map_err(|_| corrupted("Tcg length exceeds u32::MAX"))?;
    fib.set_cmds(offset, length);
    table_stream.extend_from_slice(&data);
    Ok(())
}

/// An immutable DOC package snapshot carrying its command-bar metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Vec<u8>,
    command_bars: TransactionSnapshot,
}

impl Snapshot {
    fn new(bytes: Vec<u8>, command_bars: TransactionSnapshot) -> Self {
        Self {
            bytes,
            command_bars,
        }
    }

    /// Returns the immutable semantic command-bar snapshot.
    #[must_use]
    pub fn command_bars(&self) -> &TransactionSnapshot {
        &self.command_bars
    }

    /// Returns the rendered package bytes.
    pub fn finish(&self) -> Result<Vec<u8>> {
        Ok(self.bytes.clone())
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.command_bars.edit()
    }
}

/// A package commit containing semantic and reversible OLE byte patches.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: TransactionCommit,
    package_patch: ObjectPatch,
}

impl Commit {
    /// The immutable post-edit package snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// The reversible semantic command-bar patch.
    #[must_use]
    pub fn patch(&self) -> &TransactionCommit {
        &self.patch
    }

    /// The reversible whole-CFB byte patch.
    #[must_use]
    pub fn package_patch(&self) -> &ObjectPatch {
        &self.package_patch
    }

    /// Splits the package snapshot, semantic patch, and CFB byte patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, TransactionCommit, ObjectPatch) {
        (self.snapshot, self.patch, self.package_patch)
    }
}

/// Transactional DOC package editor for the `fcCmds/lcbCmds` range.
#[derive(Debug, Clone)]
pub struct Editor {
    package: ObjectEditor,
    table_name: String,
    original: TransactionSnapshot,
    command_bars: TransactionSnapshot,
    changed: bool,
}

impl Editor {
    /// Opens an unencrypted Word 97+ package from CFB bytes.
    pub fn open(bytes: Vec<u8>) -> Result<Self> {
        Self::open_with_limits(bytes, Limits::default())
    }

    /// Opens a package with an explicit bounded OLE resource profile.
    pub fn open_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let package =
            ObjectEditor::open(bytes, Targets::new([]).map_err(PackageError::from)?, limits)
                .map_err(PackageError::from)?;
        let word = package
            .stream(&["WordDocument".to_owned()])
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?;
        let fib = FileInformationBlock::parse(word)?;
        validation::package_fib(&fib)?;
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table = package
            .stream(&[table_name.to_owned()])
            .ok_or_else(|| PackageError::StreamNotFound(table_name.to_owned()))?;
        let command_bars = TransactionSnapshot::new(parse(&fib, table)?)?;
        Ok(Self {
            package,
            table_name: table_name.to_owned(),
            original: command_bars.clone(),
            command_bars,
            changed: false,
        })
    }

    /// Returns the current immutable semantic command-bar snapshot.
    #[must_use]
    pub fn command_bars(&self) -> &TransactionSnapshot {
        &self.command_bars
    }

    /// Whether this editor has published a changed package state.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Starts an independent semantic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.command_bars.edit()
    }

    /// Sets or replaces the complete command-bar range atomically.
    pub fn set(&mut self, command_bars: CommandBars<'_>) -> Result<Commit> {
        self.install(TransactionSnapshot::new(Some(command_bars))?)
    }

    /// Replaces existing command-bar metadata, rejecting an absent range.
    pub fn replace(&mut self, command_bars: CommandBars<'_>) -> Result<Commit> {
        if !self.command_bars.is_present() {
            return Err(PackageError::InvalidFormat(
                "cannot replace missing command-bar metadata".into(),
            ));
        }
        self.set(command_bars)
    }

    /// Applies and publishes a semantic transaction atomically.
    pub fn apply(&mut self, transaction: Transaction) -> Result<Commit> {
        let commit = transaction.commit().map_err(transaction_error)?;
        if commit.before() != &self.command_bars {
            return Err(PackageError::InvalidFormat(
                "command-bar transaction snapshot conflict".into(),
            ));
        }
        self.install(commit.snapshot().clone())
    }

    /// Clears the command-bar FIB range while preserving its old table bytes.
    pub fn clear(&mut self) -> Result<Commit> {
        let mut transaction = self.edit();
        transaction.clear();
        self.apply(transaction)
    }

    /// Captures the current package as an immutable snapshot.
    pub fn snapshot(&self) -> Result<Snapshot> {
        let bytes = self.package.clone().finish().map_err(PackageError::from)?;
        Ok(Snapshot::new(bytes, self.command_bars.clone()))
    }

    /// Finishes the edit and returns rendered DOC bytes.
    pub fn finish(self) -> Result<Vec<u8>> {
        self.package.finish().map_err(PackageError::from)
    }

    /// Commits the package as an immutable snapshot with reversible patches.
    pub fn commit(self) -> Result<Commit> {
        let patch = TransactionCommit::new(self.original, self.command_bars.clone());
        let object_commit = self.package.commit().map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let package_patch = object_commit.into_patch();
        Ok(Commit {
            snapshot: Snapshot::new(bytes, self.command_bars),
            patch,
            package_patch,
        })
    }

    fn install(&mut self, snapshot: TransactionSnapshot) -> Result<Commit> {
        let patch = TransactionCommit::new(self.command_bars.clone(), snapshot.clone());
        if patch.is_noop() {
            let package_patch = self
                .package
                .clone()
                .commit()
                .map_err(PackageError::from)?
                .into_patch();
            return Ok(Commit {
                snapshot: self.snapshot()?,
                patch,
                package_patch,
            });
        }

        let mut candidate = self.clone();
        candidate.write_snapshot(&snapshot)?;
        candidate.command_bars = snapshot;
        candidate.changed = true;
        let object_commit = candidate
            .package
            .clone()
            .commit()
            .map_err(PackageError::from)?;
        let bytes = object_commit.patch().after().to_vec();
        let package_patch = object_commit.into_patch();
        let package_snapshot = Snapshot::new(bytes, candidate.command_bars.clone());
        *self = candidate;
        Ok(Commit {
            snapshot: package_snapshot,
            patch,
            package_patch,
        })
    }

    fn write_snapshot(&mut self, snapshot: &TransactionSnapshot) -> Result<()> {
        let word_path = vec!["WordDocument".to_owned()];
        let table_path = vec![self.table_name.clone()];
        let mut word = self
            .package
            .stream(&word_path)
            .ok_or_else(|| PackageError::StreamNotFound("WordDocument".into()))?
            .to_vec();
        let mut table = self
            .package
            .stream(&table_path)
            .ok_or_else(|| PackageError::StreamNotFound(self.table_name.clone()))?
            .to_vec();
        let fib = FileInformationBlock::parse(&word)?;
        let pointer = validation::pointer_location(&fib)?;
        if let Some(command_bars) = snapshot.command_bars() {
            let payload = codec::to_bytes(command_bars)?;
            let offset = u32::try_from(table.len())
                .map_err(|_| PackageError::Corrupted("table stream exceeds u32::MAX".into()))?;
            let length = u32::try_from(payload.len()).map_err(|_| {
                PackageError::Corrupted("command-bar payload exceeds u32::MAX".into())
            })?;
            table.extend_from_slice(&payload);
            word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
            word[pointer + 4..pointer + 8].copy_from_slice(&length.to_le_bytes());
        } else {
            word[pointer..pointer + 8].fill(0);
        }

        let reparsed_fib = FileInformationBlock::parse(&word)?;
        let reparsed = TransactionSnapshot::new(parse(&reparsed_fib, &table)?)?;
        if &reparsed != snapshot {
            return Err(PackageError::Corrupted(
                "command-bar snapshot failed FIB/table-stream round-trip validation".into(),
            ));
        }

        let mut package = self.package.clone();
        package
            .put_stream(&table_path, table)
            .map_err(PackageError::from)?;
        package
            .put_stream(&word_path, word)
            .map_err(PackageError::from)?;
        self.package = package;
        Ok(())
    }
}

fn transaction_error(error: TransactionError) -> PackageError {
    PackageError::InvalidFormat(format!("command-bar transaction failed: {error}"))
}

fn corrupted(message: impl Into<String>) -> crate::package::Error {
    crate::package::Error::Corrupted(message.into())
}
