//! Immutable snapshots and source-checked edits for `User Names`.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use super::model::{Limits, UserEntry, UserGuid, UserNames};
use super::package::Package;
use crate::{Error, Result};

/// An immutable, source-preserving `User Names` stream snapshot.
#[derive(Clone)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    package: Arc<Package>,
    fingerprint: u64,
}

impl Snapshot {
    /// Parse a complete `User Names` stream with default resource bounds.
    /// # Errors
    ///
    /// Returns an error when framing, the stream grammar, or a typed payload
    /// violates its bounded MS-XLS contract.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse a complete stream with explicit resource bounds.
    /// # Errors
    ///
    /// Returns an error when a bound or the stream grammar is violated.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        let limits = limits.validate()?;
        let source = bytes.as_ref();
        ensure_source_bound(source, limits)?;
        let bytes = copy_bytes(source, "retaining User Names source")?;
        Self::parse_shared_with_limits(bytes, limits)
    }

    /// Parse an already shared source allocation without copying it.
    /// # Errors
    ///
    /// Returns an error when a bound or the stream grammar is violated.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self> {
        Self::parse_shared_with_limits(bytes, Limits::default())
    }

    /// Parse an already shared source allocation with explicit bounds.
    /// # Errors
    ///
    /// Returns an error when a bound or the stream grammar is violated.
    pub fn parse_shared_with_limits(bytes: Arc<[u8]>, limits: Limits) -> Result<Self> {
        let package = Arc::new(Package::parse(&bytes, limits)?);
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            package,
        })
    }

    /// Parse and bind the snapshot to the GUIDs present in `Revision Log`.
    ///
    /// The source stream remains authoritative; the GUID list only establishes
    /// the dependency closure needed for safe collection edits.
    /// # Errors
    ///
    /// Returns an error when any `UsrInfo.guid` is absent from the supplied
    /// revision-log header set.
    pub fn parse_with_revision_guids(
        bytes: impl AsRef<[u8]>,
        revision_guids: &[[u8; 16]],
    ) -> Result<Self> {
        Self::parse_with_revision_guids_and_limits(bytes, revision_guids, Limits::default())
    }

    /// Parse and bind a stream to a revision-log GUID set with explicit limits.
    /// # Errors
    ///
    /// Returns an error when a bound or the GUID dependency closure is violated.
    pub fn parse_with_revision_guids_and_limits(
        bytes: impl AsRef<[u8]>,
        revision_guids: &[[u8; 16]],
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        let source = bytes.as_ref();
        ensure_source_bound(source, limits)?;
        let bytes = copy_bytes(source, "retaining User Names source")?;
        Self::parse_shared_with_revision_guids_and_limits(bytes, revision_guids, limits)
    }

    /// Parse a shared source allocation bound to a revision-log GUID set.
    /// # Errors
    ///
    /// Returns an error when a bound or the GUID dependency closure is violated.
    pub fn parse_shared_with_revision_guids(
        bytes: Arc<[u8]>,
        revision_guids: &[[u8; 16]],
    ) -> Result<Self> {
        Self::parse_shared_with_revision_guids_and_limits(bytes, revision_guids, Limits::default())
    }

    /// Parse shared bytes with a revision-log GUID set and explicit limits.
    /// # Errors
    ///
    /// Returns an error when a bound or the GUID dependency closure is violated.
    pub fn parse_shared_with_revision_guids_and_limits(
        bytes: Arc<[u8]>,
        revision_guids: &[[u8; 16]],
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        ensure_source_bound(&bytes, limits)?;
        let guids = copy_guids(revision_guids, limits)?;
        Self::parse_shared_with_revision_guids_and_limits_and_source(bytes, guids, limits, None)
    }

    /// Parse a bounded shared User Names allocation with an already-owned
    /// Revision Log GUID set. The package owner uses this handoff after its
    /// streaming closure scan so the GUID vector is not copied a second time.
    pub(crate) fn parse_shared_with_revision_guids_arc_and_revision_log_with_limits(
        bytes: Arc<[u8]>,
        revision_guids: Arc<[UserGuid]>,
        revision_log: Arc<[u8]>,
        limits: Limits,
    ) -> Result<Self> {
        let limits = limits.validate()?;
        ensure_source_bound(&bytes, limits)?;
        ensure_revision_log_bound(&revision_log, limits)?;
        if revision_guids.len() > limits.max_revision_guids {
            return Err(Error::UnsafeEdit(format!(
                "Revision Log contains {} GUIDs; maximum is {}",
                revision_guids.len(),
                limits.max_revision_guids
            )));
        }
        Self::parse_shared_with_revision_guids_and_limits_and_source(
            bytes,
            revision_guids,
            limits,
            Some(revision_log),
        )
    }

    fn parse_shared_with_revision_guids_and_limits_and_source(
        bytes: Arc<[u8]>,
        revision_guids: Arc<[UserGuid]>,
        limits: Limits,
        revision_log_source: Option<Arc<[u8]>>,
    ) -> Result<Self> {
        let package = Arc::new(Package::parse_with_revision_guids_and_source(
            &bytes,
            limits,
            revision_guids,
            revision_log_source,
        )?);
        Ok(Self {
            fingerprint: fingerprint(&bytes),
            bytes,
            package,
        })
    }

    /// Borrow the typed user-log projection.
    #[must_use]
    pub fn model(&self) -> &UserNames {
        self.package.model()
    }

    /// Borrow users in source order.
    #[must_use]
    pub fn users(&self) -> &[UserEntry] {
        self.model().users()
    }

    /// Number of user entries declared by `CUsr.iCount`.
    #[must_use]
    pub fn user_count(&self) -> usize {
        self.model().user_count()
    }

    /// Return the exact source bytes, including all reserved fields.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Alias for [`Self::bytes`].
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        self.bytes()
    }

    /// Share the exact source allocation without copying it.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.bytes)
    }

    /// Compact source identity used by stale-source checks.
    #[must_use]
    pub const fn fingerprint(&self) -> u64 {
        self.fingerprint
    }

    /// Whether collection insertion has a proven Revision Log dependency closure.
    #[must_use]
    pub fn has_revision_guid_closure(&self) -> bool {
        self.package.revision_guids.is_some()
    }

    /// Start a detached, failure-atomic edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction::new(self.clone())
    }

    /// Publish the exact validated source stream.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("UserNamesSnapshot")
            .field("bytes", &self.bytes.len())
            .field("users", &self.user_count())
            .field("revision_guid_closure", &self.has_revision_guid_closure())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes == other.bytes
            && self.package.revision_guids == other.package.revision_guids
            && self.package.revision_log_source == other.package.revision_log_source
    }
}

impl Eq for Snapshot {}

/// A detached, source-checked transaction over user metadata and membership.
#[derive(Debug, Clone)]
pub struct Transaction {
    source: Snapshot,
    candidate: Arc<[u8]>,
    package: Arc<Package>,
}

impl Transaction {
    fn new(source: Snapshot) -> Self {
        Self {
            candidate: Arc::clone(&source.bytes),
            package: Arc::clone(&source.package),
            source,
        }
    }

    /// Borrow the immutable source snapshot used for publication checks.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        &self.source
    }

    /// Alias for [`Self::before`].
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        self.before()
    }

    /// Borrow currently staged users.
    #[must_use]
    pub fn users(&self) -> &[UserEntry] {
        self.package.model.users()
    }

    /// Borrow currently staged model metadata.
    #[must_use]
    pub fn model(&self) -> &UserNames {
        self.package.model()
    }

    /// Materialize and validate the staged candidate.
    /// # Errors
    ///
    /// Returns an error if the staged bytes no longer satisfy the stream
    /// grammar or the bound Revision Log closure.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if self.candidate.as_ref() == self.source.bytes() {
            return Ok(self.source.clone());
        }
        self.parse_candidate(Arc::clone(&self.candidate))
    }

    /// Whether staged bytes differ from the exact source stream.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate.as_ref() != self.source.bytes()
    }

    /// Change one user name while retaining its GUID, timestamp, and reserved byte.
    /// # Errors
    ///
    /// Returns an error when the index or new name is invalid.
    pub fn set_user_name(&mut self, index: usize, name: impl Into<String>) -> Result<&mut Self> {
        let current = self.users().get(index).cloned().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "User Names index {index} is outside the collection"
            ))
        })?;
        let updated = current.with_user_name(name)?;
        if updated == current {
            return Ok(self);
        }
        let candidate = self
            .package
            .replace_user(self.candidate.as_ref(), index, &updated)?;
        self.replace_candidate(candidate)?;
        Ok(self)
    }

    /// Change one user's opening timestamp.
    /// # Errors
    ///
    /// Returns an error when the index is invalid or publication fails.
    pub fn set_user_opened_at(
        &mut self,
        index: usize,
        opened_at: crate::revision_records::ShortDtr,
    ) -> Result<&mut Self> {
        let current = self.users().get(index).cloned().ok_or_else(|| {
            Error::UnsafeEdit(format!(
                "User Names index {index} is outside the collection"
            ))
        })?;
        let updated = current.with_opened_at(opened_at);
        if updated == current {
            return Ok(self);
        }
        let candidate = self
            .package
            .replace_user(self.candidate.as_ref(), index, &updated)?;
        self.replace_candidate(candidate)?;
        Ok(self)
    }

    /// Alias for [`Self::set_user_opened_at`].
    /// # Errors
    ///
    /// Returns an error when the index is invalid or publication fails.
    pub fn set_opened_at(
        &mut self,
        index: usize,
        opened_at: crate::revision_records::ShortDtr,
    ) -> Result<&mut Self> {
        self.set_user_opened_at(index, opened_at)
    }

    /// Insert a user at a checked collection position.
    /// # Errors
    ///
    /// Returns an error when the index, user identifier, GUID closure, or
    /// configured count bound is invalid.
    pub fn insert(&mut self, index: usize, user: UserEntry) -> Result<&mut Self> {
        let candidate = self
            .package
            .insert_user(self.candidate.as_ref(), index, &user)?;
        self.replace_candidate(candidate)?;
        Ok(self)
    }

    /// Append a user to the user log.
    /// # Errors
    ///
    /// Returns an error when the user cannot be inserted safely.
    pub fn push(&mut self, user: UserEntry) -> Result<&mut Self> {
        self.insert(self.users().len(), user)
    }

    /// Alias for [`Self::push`].
    /// # Errors
    ///
    /// Returns an error when the user cannot be inserted safely.
    pub fn add_user(&mut self, user: UserEntry) -> Result<&mut Self> {
        self.push(user)
    }

    /// Remove one user and return its detached semantic value.
    /// # Errors
    ///
    /// Returns an error when the index is outside the collection.
    pub fn remove(&mut self, index: usize) -> Result<UserEntry> {
        let (candidate, removed) = self.package.remove_user(self.candidate.as_ref(), index)?;
        self.replace_candidate(candidate)?;
        Ok(removed)
    }

    /// Discard staged changes and return the original source snapshot.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Validate and publish the candidate with a reversible source-checked patch.
    /// # Errors
    ///
    /// Returns an error when the staged candidate cannot be validated.
    pub fn commit(self) -> Result<Commit> {
        let Self {
            source,
            candidate,
            package,
        } = self;
        if candidate.as_ref() == source.bytes() {
            let patch = Patch::new(source.clone(), source.clone());
            return Ok(Commit {
                snapshot: source,
                patch,
            });
        }
        let snapshot = Snapshot {
            fingerprint: fingerprint(&candidate),
            bytes: candidate,
            // `replace_candidate` reparses every changed candidate before it
            // reaches commit, so transfer that validated package instead of
            // paying for a second full User Names parse at publication.
            package,
        };
        let patch = Patch::new(source, snapshot.clone());
        Ok(Commit { snapshot, patch })
    }

    fn replace_candidate(&mut self, candidate: Vec<u8>) -> Result<()> {
        let bytes = Arc::<[u8]>::from(candidate.into_boxed_slice());
        let package = if let Some(guids) = &self.package.revision_guids {
            Package::parse_with_revision_guids_and_source(
                &bytes,
                self.package.limits,
                Arc::clone(guids),
                self.package.revision_log_source.clone(),
            )?
        } else {
            Package::parse(&bytes, self.package.limits)?
        };
        self.candidate = bytes;
        self.package = Arc::new(package);
        Ok(())
    }

    fn parse_candidate(&self, bytes: Arc<[u8]>) -> Result<Snapshot> {
        let package = if let Some(guids) = &self.package.revision_guids {
            Package::parse_with_revision_guids_and_source(
                &bytes,
                self.package.limits,
                Arc::clone(guids),
                self.package.revision_log_source.clone(),
            )?
        } else {
            Package::parse(&bytes, self.package.limits)?
        };
        Ok(Snapshot {
            fingerprint: fingerprint(&bytes),
            bytes,
            package: Arc::new(package),
        })
    }
}

/// A source-checked replacement of one complete `User Names` stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    source_fingerprint: u64,
    target_fingerprint: u64,
    source_revision_guids: Option<Arc<[UserGuid]>>,
    target_revision_guids: Option<Arc<[UserGuid]>>,
    source_revision_log_source: Option<Arc<[u8]>>,
    target_revision_log_source: Option<Arc<[u8]>>,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self {
            source_fingerprint: before.fingerprint,
            target_fingerprint: after.fingerprint,
            source_revision_guids: before.package.revision_guids.clone(),
            target_revision_guids: after.package.revision_guids.clone(),
            source_revision_log_source: before.package.revision_log_source.clone(),
            target_revision_log_source: after.package.revision_log_source.clone(),
            before: before.bytes_shared(),
            after: after.bytes_shared(),
        }
    }

    /// Fingerprint required of the source snapshot.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Fingerprint produced by applying this patch.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Exact source bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether this patch is an exact byte-for-byte no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Apply this patch only to its exact source snapshot.
    /// # Errors
    ///
    /// Returns an error when source bytes or their compact identity differ.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.fingerprint != self.source_fingerprint
            || source.bytes() != self.before()
            || source.package.revision_guids != self.source_revision_guids
            || source.package.revision_log_source != self.source_revision_log_source
        {
            return Err(Error::UnsafeEdit(
                "User Names patch source does not match its base snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(source.clone())
        } else if let Some(guids) = &source.package.revision_guids {
            Snapshot::parse_shared_with_revision_guids_and_limits_and_source(
                Arc::clone(&self.after),
                Arc::clone(guids),
                source.package.limits,
                source.package.revision_log_source.clone(),
            )
        } else {
            Snapshot::parse_shared_with_limits(Arc::clone(&self.after), source.package.limits)
        }
    }

    /// Apply the inverse replacement to the committed target.
    /// # Errors
    ///
    /// Returns an error when target bytes or their compact identity differ.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.fingerprint != self.target_fingerprint
            || target.bytes() != self.after()
            || target.package.revision_guids != self.target_revision_guids
            || target.package.revision_log_source != self.target_revision_log_source
        {
            return Err(Error::UnsafeEdit(
                "User Names patch target does not match its committed snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            Ok(target.clone())
        } else if let Some(guids) = &target.package.revision_guids {
            Snapshot::parse_shared_with_revision_guids_and_limits_and_source(
                Arc::clone(&self.before),
                Arc::clone(guids),
                target.package.limits,
                target.package.revision_log_source.clone(),
            )
        } else {
            Snapshot::parse_shared_with_limits(Arc::clone(&self.before), target.package.limits)
        }
    }

    /// Return the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            source_revision_guids: self.target_revision_guids.clone(),
            target_revision_guids: self.source_revision_guids.clone(),
            source_revision_log_source: self.target_revision_log_source.clone(),
            target_revision_log_source: self.source_revision_log_source.clone(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

/// A successful edit publication containing its snapshot and reversible patch.
#[derive(Clone, Debug, PartialEq, Eq)]
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

    /// Resulting immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible source-checked patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the publication into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the publication into its patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consume the publication into exact stream bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.finish()
    }

    /// Consume the publication into both artifacts.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

fn copy_bytes(bytes: &[u8], context: &'static str) -> Result<Arc<[u8]>> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(bytes.len())
        .map_err(|_error| Error::Allocation(context))?;
    owned.extend_from_slice(bytes);
    Ok(Arc::from(owned.into_boxed_slice()))
}

fn ensure_source_bound(bytes: &[u8], limits: Limits) -> Result<()> {
    if bytes.len() > limits.max_stream_bytes {
        return Err(Error::UnsafeEdit(format!(
            "User Names stream has {} bytes; maximum is {}",
            bytes.len(),
            limits.max_stream_bytes
        )));
    }
    Ok(())
}

fn ensure_revision_log_bound(bytes: &[u8], limits: Limits) -> Result<()> {
    if bytes.len() > limits.max_stream_bytes {
        return Err(Error::UnsafeEdit(format!(
            "Revision Log stream has {} bytes; maximum is {}",
            bytes.len(),
            limits.max_stream_bytes
        )));
    }
    Ok(())
}

fn copy_guids(guids: &[[u8; 16]], limits: Limits) -> Result<Arc<[UserGuid]>> {
    if guids.len() > limits.max_revision_guids {
        return Err(Error::UnsafeEdit(format!(
            "Revision Log contains {} GUIDs; maximum is {}",
            guids.len(),
            limits.max_revision_guids
        )));
    }
    let byte_count = guids
        .len()
        .checked_mul(size_of::<UserGuid>())
        .ok_or(Error::Allocation("sizing Revision Log GUID closure"))?;
    if byte_count > limits.max_stream_bytes {
        return Err(Error::UnsafeEdit(format!(
            "Revision Log GUID closure requires {byte_count} bytes; maximum is {}",
            limits.max_stream_bytes
        )));
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(guids.len())
        .map_err(|_error| Error::Allocation("retaining Revision Log GUIDs"))?;
    owned.extend_from_slice(guids);
    Ok(Arc::from(owned.into_boxed_slice()))
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut value = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
    }
    value
}
