//! Source-bound publication of an XLS package containing `User Names`.
//!
//! [`Snapshot`] owns the complete validated CFB artifact, while the common
//! OLE package editor owns rendering and directory metadata preservation.  A
//! transaction changes only the typed `User Names` stream; the resulting CFB
//! is reopened through the normal XLS reader and the Revision Log dependency
//! check before a commit is returned.

use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_ole_common::object::{Editor as PackageEditor, Limits as PackageLimits, Targets};

use super::{
    Limits as UserNamesLimits, Snapshot as UserNamesSnapshot, Transaction as UserNamesTransaction,
};
use crate::{Error, Result};

/// An immutable, cheaply cloned source-bound XLS package snapshot.
#[derive(Clone)]
pub struct Snapshot {
    inner: Arc<Inner>,
}

struct Inner {
    bytes: Arc<[u8]>,
    user_names_path: Vec<String>,
    user_names: UserNamesSnapshot,
    user_names_limits: UserNamesLimits,
}

impl Snapshot {
    /// Open an XLS package that contains one root `User Names` stream.
    ///
    /// The common CFB owner rejects signed, encrypted, and DRM containers.
    /// BIFF `FILEPASS` records are rejected by the complete XLS open as well;
    /// this owner never attempts to decrypt or re-encrypt a source.  All CFB
    /// streams and directory metadata are captured before an edit is exposed.
    ///
    /// # Errors
    ///
    /// Returns a CFB, workbook, User Names grammar, Revision Log closure, or
    /// protected-source error.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, UserNamesLimits::default())
    }

    /// Open an XLS package with an explicit `User Names` resource bound.
    ///
    /// The byte bound is checked against the User Names and required Revision
    /// Log directory stream lengths before the common package owner captures
    /// stream payloads. The User Names owner then applies bounded framed
    /// record and `RRDHead` closure limits while scanning Revision Log bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bound, package, workbook, User Names stream,
    /// or Revision Log dependency closure is invalid.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: UserNamesLimits) -> Result<Self> {
        validate_user_names_limits(limits)?;
        let expected_path = preflight_user_names(&bytes, limits)?;
        let (package, user_names_path) = open_package(bytes)?;
        // The preflight and the common owner inspect the same immutable input;
        // keep the exact spelling selected by the owner for the replacement.
        if user_names_path != expected_path {
            return Err(Error::UnsafeEdit(
                "User Names stream path changed between CFB validation passes".to_string(),
            ));
        }
        let bytes = Arc::<[u8]>::from(package.finish()?.into_boxed_slice());
        let user_names = parse_user_names(&bytes, limits)?;
        Ok(Self {
            inner: Arc::new(Inner {
                bytes,
                user_names_path,
                user_names,
                user_names_limits: limits,
            }),
        })
    }

    /// Exact source CFB bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.inner.bytes
    }

    /// Share the exact source CFB allocation without copying it.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.inner.bytes)
    }

    /// Typed `User Names` stream projection bound to the source Revision Log.
    #[must_use]
    pub fn user_names(&self) -> &UserNamesSnapshot {
        &self.inner.user_names
    }

    /// Alias emphasizing the typed stream owner.
    #[must_use]
    pub fn user_names_snapshot(&self) -> &UserNamesSnapshot {
        self.user_names()
    }

    /// Start a detached source-bound User Names transaction.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        self.edit_user_names()
    }

    /// Start a detached source-bound User Names transaction.
    #[must_use]
    pub fn edit_user_names(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            user_names: self.inner.user_names.edit(),
        }
    }

    /// Return the exact validated source artifact.
    #[must_use]
    pub fn finish(&self) -> Vec<u8> {
        self.inner.bytes.to_vec()
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("UserNamesPackageSnapshot")
            .field("artifact_bytes", &self.bytes().len())
            .field("user_count", &self.user_names().user_count())
            .finish()
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.bytes() == other.bytes()
    }
}

impl Eq for Snapshot {}

/// A detached source-bound transaction over one complete XLS package.
#[derive(Clone)]
pub struct Transaction {
    source: Snapshot,
    user_names: UserNamesTransaction,
}

impl Transaction {
    /// The immutable package snapshot used as the publication source.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// The immutable package snapshot used as the publication source.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        self.source()
    }

    /// Borrow staged User Names entries in source order.
    #[must_use]
    pub fn users(&self) -> &[super::UserEntry] {
        self.user_names.users()
    }

    /// Borrow the staged User Names semantic projection.
    #[must_use]
    pub fn model(&self) -> &super::UserNames {
        self.user_names.model()
    }

    /// Whether the staged User Names stream differs from its source bytes.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.user_names.is_changed()
    }

    /// Stage a User Names display-name change.
    ///
    /// # Errors
    ///
    /// Returns an index, name, or bounded stream error without changing this
    /// transaction.
    pub fn set_user_name(&mut self, index: usize, name: impl Into<String>) -> Result<&mut Self> {
        self.user_names.set_user_name(index, name)?;
        Ok(self)
    }

    /// Stage a User Names opening-time change.
    ///
    /// # Errors
    ///
    /// Returns an index or bounded stream error without changing this
    /// transaction.
    pub fn set_user_opened_at(
        &mut self,
        index: usize,
        opened_at: crate::revision_records::ShortDtr,
    ) -> Result<&mut Self> {
        self.user_names.set_user_opened_at(index, opened_at)?;
        Ok(self)
    }

    /// Alias for [`Self::set_user_opened_at`].
    ///
    /// # Errors
    ///
    /// Returns an index or bounded stream error without changing this
    /// transaction.
    pub fn set_opened_at(
        &mut self,
        index: usize,
        opened_at: crate::revision_records::ShortDtr,
    ) -> Result<&mut Self> {
        self.set_user_opened_at(index, opened_at)
    }

    /// Stage insertion of a User Names entry.
    ///
    /// # Errors
    ///
    /// Returns an index, GUID-closure, duplicate-ID, or bound error without
    /// changing this transaction.
    pub fn insert(&mut self, index: usize, user: super::UserEntry) -> Result<&mut Self> {
        self.user_names.insert(index, user)?;
        Ok(self)
    }

    /// Append a User Names entry.
    ///
    /// # Errors
    ///
    /// Returns a GUID-closure, duplicate-ID, or bound error without changing
    /// this transaction.
    pub fn push(&mut self, user: super::UserEntry) -> Result<&mut Self> {
        self.user_names.push(user)?;
        Ok(self)
    }

    /// Alias for [`Self::push`].
    ///
    /// # Errors
    ///
    /// Returns a GUID-closure, duplicate-ID, or bound error without changing
    /// this transaction.
    pub fn add_user(&mut self, user: super::UserEntry) -> Result<&mut Self> {
        self.push(user)
    }

    /// Remove a User Names entry and return its detached semantic value.
    ///
    /// # Errors
    ///
    /// Returns an index or bounded stream error without changing the source
    /// package.
    pub fn remove(&mut self, index: usize) -> Result<super::UserEntry> {
        self.user_names.remove(index)
    }

    /// Return the source package and discard staged changes.
    #[must_use]
    pub fn rollback(self) -> Snapshot {
        self.source
    }

    /// Reparse the staged User Names stream as a complete package snapshot.
    ///
    /// This renders and reopens the candidate through the common CFB owner;
    /// it is useful for callers that need a read view before final commit.
    ///
    /// # Errors
    ///
    /// Returns a CFB, workbook, stream, or dependency-closure error.
    pub fn snapshot(&self) -> Result<Snapshot> {
        self.clone().commit().map(|commit| commit.snapshot)
    }

    /// Validate and publish the complete XLS package with a reversible patch.
    ///
    /// The common publication owner retains untouched stream payloads and CFB
    /// directory metadata.  The candidate is reopened through the complete
    /// XLS facade, including the current Revision Log GUID closure, before it
    /// is returned.
    ///
    /// # Errors
    ///
    /// Returns a CFB, workbook, stream, User Names, or dependency-closure
    /// error without publishing partial bytes.
    pub fn commit(self) -> Result<Commit> {
        let Self { source, user_names } = self;
        let user_names_commit = user_names.commit()?;
        if !user_names_commit.changed() {
            return Ok(Commit {
                patch: Patch::new(source.bytes_shared(), source.bytes_shared()),
                snapshot: source,
            });
        }

        let replacement =
            Arc::<[u8]>::from(user_names_commit.snapshot().finish().into_boxed_slice());
        let mut package = PackageEditor::open(
            source.bytes().to_vec(),
            Targets::default(),
            PackageLimits::default(),
        )?;
        let rendered =
            package.put_stream_shared_with_rendered(&source.inner.user_names_path, replacement)?;
        let snapshot = Snapshot::from_bytes_with_limits(rendered, source.inner.user_names_limits)?;
        let patch = Patch::new(source.bytes_shared(), snapshot.bytes_shared());
        Ok(Commit { snapshot, patch })
    }
}

impl fmt::Debug for Transaction {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("UserNamesPackageTransaction")
            .field("source", &self.source)
            .field("changed", &self.is_changed())
            .finish()
    }
}

/// A successful complete-CFB User Names publication.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Reopened target package snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Exact-source reversible package patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Whether the complete CFB artifact changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.is_noop()
    }

    /// Consume the publication into its target snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Consume the publication into its reversible patch.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Consume the publication into exact complete-CFB bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.snapshot.finish()
    }

    /// Consume the publication into its target snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// An exact-source reversible replacement of a complete XLS CFB artifact.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    fn new(before: Arc<[u8]>, after: Arc<[u8]>) -> Self {
        Self { before, after }
    }

    /// Exact source CFB bytes required by this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact target CFB bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the complete CFB artifact is byte-for-byte unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Alias for [`Self::is_noop`].
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.is_noop()
    }

    /// Apply this patch only to its exact source package snapshot.
    ///
    /// # Errors
    ///
    /// Returns a stale-source error or rejects the target during complete
    /// package reopen.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::UnsafeEdit(
                "User Names package patch source does not match its base snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes_with_limits(self.after.to_vec(), source.inner.user_names_limits)
    }

    /// Apply the inverse replacement to its exact committed target.
    ///
    /// # Errors
    ///
    /// Returns a stale-target error or rejects the source during complete
    /// package reopen.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.bytes() != self.after() {
            return Err(Error::UnsafeEdit(
                "User Names package patch target does not match its committed snapshot".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(target.clone());
        }
        Snapshot::from_bytes_with_limits(self.before.to_vec(), target.inner.user_names_limits)
    }

    /// Return the exact inverse replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

fn validate_user_names_limits(limits: UserNamesLimits) -> Result<()> {
    limits.validate().map(|_| ())
}

fn preflight_user_names(bytes: &[u8], limits: UserNamesLimits) -> Result<Vec<String>> {
    let ole = OleFile::open(Cursor::new(bytes))?;
    let path = find_user_names_path(&ole)?;
    let maximum = u64::try_from(limits.max_stream_bytes).unwrap_or(u64::MAX);
    let revision_path = find_revision_log_path(&ole)?;
    let revision_len = ole.stream_len(&[revision_path[0].as_str()])?;
    if revision_len > maximum {
        return Err(Error::UnsafeEdit(format!(
            "Revision Log stream has {revision_len} bytes; maximum is {maximum}"
        )));
    }
    let stream_len = ole.stream_len(&[path[0].as_str()])?;
    if stream_len > maximum {
        return Err(Error::UnsafeEdit(format!(
            "User Names stream has {stream_len} bytes; maximum is {maximum}"
        )));
    }
    Ok(path)
}

fn open_package(bytes: Vec<u8>) -> Result<(PackageEditor, Vec<String>)> {
    let (package, ole) =
        PackageEditor::open_with_ole_file(bytes, Targets::default(), PackageLimits::default())?;
    let path = find_user_names_path(&ole)?;
    Ok((package, path))
}

fn find_user_names_path<R: std::io::Read + std::io::Seek>(ole: &OleFile<R>) -> Result<Vec<String>> {
    let mut found = None;
    for path in ole.list_streams() {
        if path.len() != 1 || !path[0].eq_ignore_ascii_case(super::USER_NAMES_STREAM_NAME) {
            continue;
        }
        if found.is_some() {
            return Err(Error::InvalidData(
                "CFB contains more than one User Names stream".to_string(),
            ));
        }
        found = Some(path);
    }
    found.ok_or_else(|| Error::InvalidData("XLS package has no root User Names stream".to_string()))
}

fn find_revision_log_path<R: std::io::Read + std::io::Seek>(
    ole: &OleFile<R>,
) -> Result<Vec<String>> {
    let mut found = None;
    for path in ole.list_streams() {
        if path.len() != 1
            || !path[0].eq_ignore_ascii_case(crate::revision_log::REVISION_LOG_STREAM_NAME)
        {
            continue;
        }
        if found.is_some() {
            return Err(Error::InvalidData(
                "CFB contains more than one Revision Log stream".to_string(),
            ));
        }
        found = Some(path);
    }
    found.ok_or_else(|| {
        Error::InvalidData("User Names stream requires a Revision Log stream".to_string())
    })
}

fn parse_user_names(bytes: &[u8], limits: UserNamesLimits) -> Result<UserNamesSnapshot> {
    let mut workbook = crate::Workbook::new(Cursor::new(bytes))?;
    workbook
        .user_names_with_limits(limits)?
        .ok_or_else(|| Error::InvalidData("XLS package has no User Names stream".to_string()))
}
