//! Source-preserving publication for BIFF8 shared-workbook revision logs.
//!
//! The revision stream is intentionally edited as inert metadata.  This
//! layer never interprets undo payloads, evaluates formulas, resolves
//! conflicts, or acquires any of the locks represented by the stream.  A
//! transaction changes only fixed-width fields whose wire locations are
//! defined by [MS-XLS]; every other byte remains owned by the source snapshot.

use std::sync::Arc;

use crate::{Error, Result};

use super::RevisionRecordHeader;
use super::codec::{
    RECORD_HEADER_LEN, RR_INSERT_SH_RECORD_TYPE, RRD_CHG_CELL_RECORD_TYPE,
    RRD_CONFLICT_RECORD_TYPE, RRD_HEAD_MAX_USER_CHARS, RRD_HEAD_RECORD_TYPE,
    RRD_HEAD_USER_FIELD_LEN, RRD_INS_DEL_RECORD_TYPE, RRD_LEN, RRD_MOVE_RECORD_TYPE,
    RRD_REN_SHEET_RECORD_TYPE, RRD_USER_VIEW_RECORD_TYPE,
};
use crate::revision_log::{RevisionLog, parse_revision_log_stream};

const FLAGS_OFFSET_IN_RRD: usize = 10;
const RRD_HEAD_USER_COUNT_OFFSET: usize = RRD_LEN + 16 + 2;
const RRD_HEAD_USER_FIELD_OFFSET: usize = RRD_HEAD_USER_COUNT_OFFSET + 2;

/// The inert review state carried by the accepted and undo bits of an RRD.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionFlags {
    accepted: bool,
    undo_action: bool,
}

impl RevisionFlags {
    /// Creates a review state without changing any other RRD flag.
    #[must_use]
    pub const fn new(accepted: bool, undo_action: bool) -> Self {
        Self {
            accepted,
            undo_action,
        }
    }

    /// Whether the revision has been reviewed and accepted.
    #[must_use]
    pub const fn accepted(self) -> bool {
        self.accepted
    }

    /// Whether the revision was created by an undo action.
    #[must_use]
    pub const fn undo_action(self) -> bool {
        self.undo_action
    }

    /// Returns this state with a new accepted bit.
    #[must_use]
    pub const fn with_accepted(self, accepted: bool) -> Self {
        Self { accepted, ..self }
    }

    /// Returns this state with a new undo-action bit.
    #[must_use]
    pub const fn with_undo_action(self, undo_action: bool) -> Self {
        Self {
            undo_action,
            ..self
        }
    }
}

/// An immutable, source-preserving Revision Log stream.
///
/// The parsed [`RevisionLog`] is the semantic view.  The exact stream bytes
/// are retained beside it so a no-op publication does not normalize BIFF
/// records, padding, or opaque revision payloads.
#[derive(Debug, Clone)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    log: RevisionLog,
    fingerprint: u64,
}

impl Snapshot {
    /// Parses and retains a validated BIFF8 `Revision Log` stream.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        let bytes = Arc::<[u8]>::from(bytes.as_ref().to_vec().into_boxed_slice());
        Self::parse_shared(bytes)
    }

    /// Parses a shared byte allocation without copying it.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self> {
        let log = parse_revision_log_stream(&bytes)?;
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            log,
        })
    }

    /// Returns the typed, inert revision-log view.
    #[must_use]
    pub const fn log(&self) -> &RevisionLog {
        &self.log
    }

    /// Alias for [`Self::log`] for callers working with the stream name.
    #[must_use]
    pub const fn revision_log(&self) -> &RevisionLog {
        self.log()
    }

    /// Returns the exact source stream bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Returns shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// Returns the compact source identity used for stale checks.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.fingerprint
    }

    /// Starts a detached, failure-atomic transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Publishes the exact validated source stream.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
    }
}

impl Eq for Snapshot {}

/// A failure-atomic edit over inert Revision Log metadata.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Vec<u8>,
}

impl Transaction {
    fn new(source: Snapshot) -> Self {
        Self {
            candidate: source.bytes().to_vec(),
            source,
        }
    }

    /// Returns the immutable source snapshot used by this transaction.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Parses the current candidate without publishing it.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_slice() == self.source.bytes() {
            Ok(self.source.clone())
        } else {
            Snapshot::parse(&self.candidate)
        }
    }

    /// Whether a staged edit changes any stream byte.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.as_slice() != self.source.bytes()
    }

    /// Replaces the accepted and undo-action bits of one revision.
    ///
    /// `revision_id` is unique within a valid shared-workbook revision
    /// sequence.  If a producer violates that identity rule, the edit is
    /// rejected as ambiguous instead of changing more than one record.
    pub fn set_revision_flags(
        &mut self,
        revision_id: i32,
        flags: RevisionFlags,
    ) -> Result<&mut Self> {
        if revision_id < 0 {
            return Err(Error::UnsafeEdit(
                "revision identifiers cannot be negative".to_string(),
            ));
        }
        let offset = find_revision_offset(&self.candidate, revision_id)?;
        let mut candidate = self.candidate.clone();
        let current = u16::from_le_bytes([
            candidate[offset + FLAGS_OFFSET_IN_RRD],
            candidate[offset + FLAGS_OFFSET_IN_RRD + 1],
        ]);
        let mut updated = current & !0x0003;
        if flags.accepted() {
            updated |= 0x0001;
        }
        if flags.undo_action() {
            updated |= 0x0002;
        }
        if current == updated {
            return Ok(self);
        }
        candidate[offset + FLAGS_OFFSET_IN_RRD..offset + FLAGS_OFFSET_IN_RRD + 2]
            .copy_from_slice(&updated.to_le_bytes());
        self.replace_candidate(candidate)?;
        Ok(self)
    }

    /// Updates only the accepted bit of one revision.
    pub fn set_revision_accepted(&mut self, revision_id: i32, accepted: bool) -> Result<&mut Self> {
        let current = revision_flags(&self.candidate, revision_id)?;
        self.set_revision_flags(revision_id, current.with_accepted(accepted))
    }

    /// Updates only the undo-action bit of one revision.
    pub fn set_revision_undo_action(
        &mut self,
        revision_id: i32,
        undo_action: bool,
    ) -> Result<&mut Self> {
        let current = revision_flags(&self.candidate, revision_id)?;
        self.set_revision_flags(revision_id, current.with_undo_action(undo_action))
    }

    /// Replaces one `RRDHead.stUser` value while retaining its fixed wire
    /// envelope, timestamp, GUID, and all other headers.
    pub fn set_header_user_name(
        &mut self,
        header_index: usize,
        user_name: &str,
    ) -> Result<&mut Self> {
        let (payload_offset, payload) = find_header(&self.candidate, header_index)?;
        let head = super::RrdHead::parse_payload(payload)?;
        if head.user_name() == user_name {
            return Ok(self);
        }
        let mut candidate = self.candidate.clone();
        let cch_offset = payload_offset + RRD_HEAD_USER_COUNT_OFFSET;
        let field_offset = payload_offset + RRD_HEAD_USER_FIELD_OFFSET;
        let field_end = field_offset + RRD_HEAD_USER_FIELD_LEN;
        let cch = encode_user_name(&mut candidate[field_offset..field_end], user_name)?;
        candidate[cch_offset..cch_offset + 2].copy_from_slice(&cch.to_le_bytes());
        self.replace_candidate(candidate)?;
        Ok(self)
    }

    /// Discards all staged edits and returns the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validates and publishes the candidate with a reversible source-checked
    /// patch.  Failed validation leaves the transaction's candidate intact.
    pub fn commit(self) -> Result<Commit> {
        let source = self.source;
        if self.candidate.as_slice() == source.bytes() {
            let patch = Patch::new(source.clone(), source.clone());
            return Ok(Commit {
                snapshot: source,
                patch,
            });
        }
        let snapshot = Snapshot::parse(self.candidate)?;
        let patch = Patch::new(source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    fn replace_candidate(&mut self, candidate: Vec<u8>) -> Result<()> {
        // Parse before publishing the candidate.  This gives each operation
        // failure atomicity and keeps the transaction's existing bytes intact
        // when a future producer-specific invariant is violated.
        Snapshot::parse(&candidate)?;
        self.candidate = candidate;
        Ok(())
    }
}

/// A successful publication containing the typed result and its reversible
/// source-checked patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Whether publication changed any stream byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Returns the post-edit snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Returns the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes the publication into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consumes the publication into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consumes the publication into exact stream bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.finish()
    }
}

/// A source-checked, reversible replacement of one complete Revision Log
/// stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            source_fingerprint: before.fingerprint(),
            target_fingerprint: after.fingerprint(),
            before: before.bytes_shared(),
            after: after.bytes_shared(),
        }
    }

    /// Returns the source fingerprint required by this patch.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Returns the target fingerprint produced by this patch.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Returns the exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Returns the exact bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Applies the patch only to its exact source snapshot.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint() != self.source_fingerprint || source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "Revision Log patch source does not match its base snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(source.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.after))
        }
    }

    /// Applies the exact inverse replacement to its target snapshot.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.fingerprint() != self.target_fingerprint || target.bytes() != self.after() {
            return Err(Error::UnsafeEdit(
                "Revision Log patch target does not match its committed snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(target.clone())
        } else {
            Snapshot::parse_shared(Arc::clone(&self.before))
        }
    }

    /// Returns the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}

fn editable_record_type(record_type: u16) -> bool {
    matches!(
        record_type,
        RRD_REN_SHEET_RECORD_TYPE
            | RRD_INS_DEL_RECORD_TYPE
            | RRD_MOVE_RECORD_TYPE
            | RRD_CHG_CELL_RECORD_TYPE
            | RRD_CONFLICT_RECORD_TYPE
            | RR_INSERT_SH_RECORD_TYPE
            | RRD_USER_VIEW_RECORD_TYPE
    )
}

fn find_revision_offset(bytes: &[u8], revision_id: i32) -> Result<usize> {
    let mut offset = 0usize;
    let mut found = None;
    while offset < bytes.len() {
        let (record_type, payload_start, payload_end) = frame(bytes, offset)?;
        let payload = &bytes[payload_start..payload_end];
        if editable_record_type(record_type) && payload.len() >= RRD_LEN {
            let header = RevisionRecordHeader::parse(record_type, payload, false)?;
            if header.revision_id() == revision_id {
                if found.is_some() {
                    return Err(Error::UnsafeEdit(format!(
                        "revision identifier {revision_id} is ambiguous"
                    )));
                }
                found = Some(payload_start);
            }
        }
        offset = payload_end;
    }
    found.ok_or_else(|| {
        Error::UnsafeEdit(format!(
            "revision identifier {revision_id} was not found in the Revision Log"
        ))
    })
}

fn revision_flags(bytes: &[u8], revision_id: i32) -> Result<RevisionFlags> {
    let offset = find_revision_offset(bytes, revision_id)?;
    let flags = u16::from_le_bytes([
        bytes[offset + FLAGS_OFFSET_IN_RRD],
        bytes[offset + FLAGS_OFFSET_IN_RRD + 1],
    ]);
    Ok(RevisionFlags::new(flags & 0x0001 != 0, flags & 0x0002 != 0))
}

fn find_header(bytes: &[u8], header_index: usize) -> Result<(usize, &[u8])> {
    let mut offset = 0usize;
    let mut current = 0usize;
    while offset < bytes.len() {
        let (record_type, payload_start, payload_end) = frame(bytes, offset)?;
        if record_type == RRD_HEAD_RECORD_TYPE {
            if current == header_index {
                let payload = &bytes[payload_start..payload_end];
                if payload.len() < RRD_HEAD_USER_FIELD_OFFSET + RRD_HEAD_USER_FIELD_LEN {
                    return Err(Error::InvalidRecord {
                        record_type,
                        message: "RRDHead payload is shorter than its fixed user field".to_string(),
                    });
                }
                return Ok((payload_start, payload));
            }
            current += 1;
        }
        offset = payload_end;
    }
    Err(Error::UnsafeEdit(format!(
        "revision header index {header_index} is out of range"
    )))
}

fn encode_user_name(field: &mut [u8], user_name: &str) -> Result<u16> {
    if field.len() != RRD_HEAD_USER_FIELD_LEN {
        return Err(Error::InvalidLength {
            expected: RRD_HEAD_USER_FIELD_LEN,
            found: field.len(),
        });
    }
    if user_name.chars().any(|character| character == '\0') {
        return Err(Error::UnsafeEdit(
            "RRDHead user names cannot contain NUL characters".to_string(),
        ));
    }
    let utf16_len = user_name.encode_utf16().count();
    if utf16_len > RRD_HEAD_MAX_USER_CHARS {
        return Err(Error::UnsafeEdit(format!(
            "RRDHead user name has {utf16_len} UTF-16 characters; maximum is {RRD_HEAD_MAX_USER_CHARS}"
        )));
    }
    let compressed = user_name
        .chars()
        .all(|character| u32::from(character) <= 0xFF);
    let count = if compressed {
        user_name.chars().count()
    } else {
        utf16_len
    };
    if count > RRD_HEAD_MAX_USER_CHARS {
        return Err(Error::UnsafeEdit(format!(
            "RRDHead user name has {count} characters; maximum is {RRD_HEAD_MAX_USER_CHARS}"
        )));
    }
    let encoded_len = 1usize
        + count
            .checked_mul(if compressed { 1 } else { 2 })
            .ok_or(Error::Allocation("encoding RRDHead user name"))?;
    if encoded_len > field.len() {
        return Err(Error::UnsafeEdit(
            "RRDHead user name does not fit its fixed field".to_string(),
        ));
    }

    field[0] = if compressed { 0 } else { 1 };
    if compressed {
        for (index, character) in user_name.chars().enumerate() {
            field[1 + index] = character as u8;
        }
    } else {
        for (index, unit) in user_name.encode_utf16().enumerate() {
            let start = 1 + index * 2;
            field[start..start + 2].copy_from_slice(&unit.to_le_bytes());
        }
    }
    u16::try_from(count).map_err(|_| Error::Allocation("encoding RRDHead user name"))
}

fn frame(bytes: &[u8], offset: usize) -> Result<(u16, usize, usize)> {
    let header_end = offset
        .checked_add(RECORD_HEADER_LEN)
        .ok_or(Error::Allocation("framing Revision Log records"))?;
    let header = bytes.get(offset..header_end).ok_or_else(|| {
        Error::UnexpectedEndOfStream("truncated Revision Log record header".to_string())
    })?;
    let record_type = u16::from_le_bytes([header[0], header[1]]);
    let length = usize::from(u16::from_le_bytes([header[2], header[3]]));
    let payload_start = header_end;
    let payload_end = payload_start
        .checked_add(length)
        .ok_or(Error::Allocation("framing Revision Log record payload"))?;
    if payload_end > bytes.len() {
        return Err(Error::InvalidRecord {
            record_type,
            message: format!("record payload of {length} bytes is truncated"),
        });
    }
    Ok((record_type, payload_start, payload_end))
}

#[cfg(test)]
#[path = "transaction_tests.rs"]
mod tests;
