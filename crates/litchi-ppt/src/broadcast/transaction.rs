//! Source-checked semantic edits for one broadcast metadata container.

use std::sync::Arc;

use crate::package::{Error, Result};
use crate::records::Record;

use super::codec;
use super::model::{Broadcast, BroadcastProperties, UnknownRecord};

/// Deterministic identity of one exact serialized broadcast container.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64 ^ bytes.len() as u64;
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Compact source fingerprint suitable for optimistic owner checks.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// One typed replacement staged by a broadcast transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Change {
    /// Replace the complete typed broadcast projection while retaining opaque
    /// source children owned by the snapshot.
    Replace { before: Broadcast, after: Broadcast },
}

impl Change {
    fn inverse(&self) -> Self {
        match self {
            Self::Replace { before, after } => Self::Replace {
                before: after.clone(),
                after: before.clone(),
            },
        }
    }

    /// Typed projection before this operation.
    #[must_use]
    pub fn before(&self) -> &Broadcast {
        match self {
            Self::Replace { before, .. } => before,
        }
    }

    /// Typed projection after this operation.
    #[must_use]
    pub fn after(&self) -> &Broadcast {
        match self {
            Self::Replace { after, .. } => after,
        }
    }
}

/// Immutable, lossless snapshot of one complete
/// `BroadcastDocInfo9Container` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    broadcast: Broadcast,
    unknown_records: Vec<UnknownRecord>,
    revision: Revision,
}

impl Snapshot {
    /// Parse exactly one complete broadcast container and retain its source.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let source = bytes.as_ref();
        let (record, consumed) = Record::parse_strict(source, 0)?;
        if consumed != source.len() {
            return Err(Error::Corrupted(
                "broadcast snapshot contains trailing bytes".into(),
            ));
        }
        Self::from_record_and_bytes(&record, source.to_vec())
    }

    /// Capture owned source bytes without a caller-side borrow.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::parse(bytes)
    }

    /// Capture a validated typed broadcast using canonical source bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_broadcast(broadcast: &Broadcast) -> Result<Self> {
        let record = broadcast.to_record()?;
        let bytes = codec::record_bytes(
            record.version,
            record.instance,
            record.record_type_raw,
            &record.data,
        )?;
        Self::parse(bytes)
    }

    fn from_record_and_bytes(record: &Record, bytes: Vec<u8>) -> Result<Self> {
        let parsed = codec::parse_lossless(record)?;
        let shared: Arc<[u8]> = Arc::from(bytes.into_boxed_slice());
        Ok(Self {
            revision: Revision::from_bytes(&shared),
            bytes: shared,
            broadcast: parsed.value,
            unknown_records: parsed.unknown_records,
        })
    }

    /// Borrow the typed broadcast projection.
    #[must_use]
    pub const fn broadcast(&self) -> &Broadcast {
        &self.broadcast
    }

    /// Borrow source-order opaque child records.
    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    /// Exact source or committed record bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Compact identity of the exact serialized source.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an isolated semantic edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.broadcast.clone(),
            changes: Vec::new(),
        }
    }
}

/// Isolated, failure-atomic editor over one source snapshot.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Broadcast,
    changes: Vec<Change>,
}

impl Transaction {
    /// Immutable source snapshot used for stale-source checks.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Current typed broadcast projection.
    #[must_use]
    pub const fn broadcast(&self) -> &Broadcast {
        &self.candidate
    }

    /// Opaque source records retained by every candidate publication.
    #[must_use]
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        self.source.unknown_records()
    }

    /// Whether staged semantics differ from the source projection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.broadcast
    }

    /// Typed operations staged in transaction order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Replace all typed broadcast fields atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, replacement: Broadcast) -> Result<()> {
        replacement.validate()?;
        self.replace_candidate(replacement)
    }

    /// Alias for callers that prefer setter terminology.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_broadcast(&mut self, replacement: Broadcast) -> Result<()> {
        self.replace(replacement)
    }

    /// Set one optional inert title.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_title(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.title, value)
    }

    /// Set one optional inert description.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_description(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.description, value)
    }

    /// Set one optional inert speaker name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_speaker(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.speaker, value)
    }

    /// Set one optional inert contact name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_contact(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.contact, value)
    }

    /// Set one optional inert remote server name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_remote_server_name(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.remote_server_name, value)
    }

    /// Set one optional inert email address.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_email_address(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.email_address, value)
    }

    /// Set one optional inert email display name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_email_name(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.email_name, value)
    }

    /// Set one optional inert HTTP chat URL.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_chat_url(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.chat_url, value)
    }

    /// Set one optional inert archive directory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_archive_directory(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.archive_directory, value)
    }

    /// Set one optional inert `NetShow` base directory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_netshow_files_base_directory(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(
            |broadcast| &mut broadcast.netshow_files_base_directory,
            value,
        )
    }

    /// Set one optional inert `NetShow` directory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_netshow_files_directory(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.netshow_files_directory, value)
    }

    /// Set one optional inert `NetShow` server name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_netshow_server_name(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.netshow_server_name, value)
    }

    /// Set one optional inert calendar entry identifier.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_entry_id(&mut self, value: Option<String>) -> Result<()> {
        self.set_optional_string(|broadcast| &mut broadcast.entry_id, value)
    }

    /// Set the inert `PowerPoint` files base directory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_ppt_files_base_directory(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(
            |broadcast| &mut broadcast.ppt_files_base_directory,
            value.into(),
        )
    }

    /// Set the inert `PowerPoint` files directory.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_ppt_files_directory(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.ppt_files_directory, value.into())
    }

    /// Set the inert `PowerPoint` files base URL or UNC path.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_ppt_files_base_url(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.ppt_files_base_url, value.into())
    }

    /// Set the inert user-name file fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_user_name(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.user_name, value.into())
    }

    /// Set the inert broadcast date/time file fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_broadcast_date_time(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.broadcast_date_time, value.into())
    }

    /// Set the inert presentation-name file fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_presentation_name(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.presentation_name, value.into())
    }

    /// Set the inert ASD file UNC path.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_asd_file_name(&mut self, value: impl Into<String>) -> Result<()> {
        self.set_required_string(|broadcast| &mut broadcast.asd_file_name, value.into())
    }

    /// Replace the typed broadcast flags and inert timestamps atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_properties(&mut self, value: BroadcastProperties) -> Result<()> {
        let mut candidate = self.candidate.clone();
        candidate.properties = value;
        self.replace_candidate(candidate)
    }

    /// Capture the current candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot> {
        self.build_snapshot()
    }

    /// Publish the candidate atomically with a reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        let snapshot = self.build_snapshot()?;
        let changes = if snapshot.bytes == self.source.bytes {
            Vec::new()
        } else {
            self.changes
        };
        let patch = Patch {
            base: self.source.revision,
            target: snapshot.revision,
            before: self.source.bytes.clone(),
            after: snapshot.bytes.clone(),
            changes,
        };
        Ok(Commit { snapshot, patch })
    }

    /// Alias for move-owned writer terminology.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn finish(self) -> Result<Commit> {
        self.commit()
    }

    /// Discard all staged edits and recover the exact source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    fn set_optional_string(
        &mut self,
        set: impl FnOnce(&mut Broadcast) -> &mut Option<String>,
        value: Option<String>,
    ) -> Result<()> {
        let mut candidate = self.candidate.clone();
        *set(&mut candidate) = value;
        self.replace_candidate(candidate)
    }

    fn set_required_string(
        &mut self,
        set: impl FnOnce(&mut Broadcast) -> &mut String,
        value: String,
    ) -> Result<()> {
        let mut candidate = self.candidate.clone();
        *set(&mut candidate) = value;
        self.replace_candidate(candidate)
    }

    fn replace_candidate(&mut self, candidate: Broadcast) -> Result<()> {
        candidate.validate()?;
        if candidate == self.candidate {
            return Ok(());
        }
        let before = std::mem::replace(&mut self.candidate, candidate.clone());
        self.changes.push(Change::Replace {
            before,
            after: candidate,
        });
        Ok(())
    }

    fn build_snapshot(&self) -> Result<Snapshot> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        let record = self
            .candidate
            .to_record_lossless(&self.source.unknown_records)?;
        let bytes = codec::record_bytes(
            record.version,
            record.instance,
            record.record_type_raw,
            &record.data,
        )?;
        Snapshot::parse(bytes)
    }
}

/// A successful immutable target and its source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Published target snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible patch from source to target.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Published target typed broadcast projection.
    #[must_use]
    pub const fn broadcast(&self) -> &Broadcast {
        self.snapshot.broadcast()
    }

    /// Undo this commit against its exact target snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.undo(current)
    }

    /// Redo this commit against its exact source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.patch.redo(current)
    }

    /// Split the commit into its target and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A source-checked reversible patch for one complete broadcast container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    changes: Vec<Change>,
}

impl Patch {
    /// Source revision required for forward application.
    #[must_use]
    pub const fn base(&self) -> Revision {
        self.base
    }

    /// Target revision produced by forward application.
    #[must_use]
    pub const fn target(&self) -> Revision {
        self.target
    }

    /// Typed operations represented by this patch.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Exact source bytes bound to this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Alias for callers that name the source bytes explicitly.
    #[must_use]
    pub fn before_bytes(&self) -> &[u8] {
        self.before()
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Alias for callers that name the target bytes explicitly.
    #[must_use]
    pub fn after_bytes(&self) -> &[u8] {
        self.after()
    }

    /// Whether this patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.as_ref() == self.after.as_ref()
    }

    /// Apply only to the exact source snapshot used to create this patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn apply(&self, current: &Snapshot) -> Result<Snapshot> {
        if current.revision != self.base || current.bytes.as_ref() != self.before.as_ref() {
            return Err(Error::InvalidFormat(
                "broadcast patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(current.clone());
        }
        Snapshot::parse(self.after.as_ref())
    }

    /// Apply the inverse to the exact committed target.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn undo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.inverse().apply(current)
    }

    /// Reapply this patch to its exact source.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn redo(&self, current: &Snapshot) -> Result<Snapshot> {
        self.apply(current)
    }

    /// Build a source-checked inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            base: self.target,
            target: self.base,
            before: self.after.clone(),
            after: self.before.clone(),
            changes: self.changes.iter().rev().map(Change::inverse).collect(),
        }
    }
}
