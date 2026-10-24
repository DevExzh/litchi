//! The XLS `User Names` edit family (record 0351 §2.3).
//!
//! [`crate::Package::user_names`] reads the shared-workbook user list as a
//! [`Snapshot`] of the complete validated CFB artifact;
//! [`crate::Package::edit_user_names`] stages typed `User Names` changes and
//! [`Edit::commit`] publishes a new [`crate::Package`] with an exact
//! [`Patch`] and [`Diagnostics`]. The common OLE package editor owns
//! rendering and directory metadata preservation; the resulting CFB is
//! reopened through the normal XLS reader and the Revision Log dependency
//! check before the commit is returned.

use litchi_core::edit::RefusalKind;
use std::fmt;
use std::io::Cursor;
use std::sync::Arc;

use litchi_cfb::OleFile;
use litchi_ole_common::object::{
    Limits as PackageLimits, OpenOptions, Package as ObjectPackage, Targets,
};

use super::{Edit as UserNamesEdit, Limits as UserNamesLimits, Snapshot as UserNamesSnapshot};
/// The User Names family error and result: the crate's.
pub use crate::{Error, Result};

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
    pub fn from_bytes(bytes: Vec<u8>, limits: UserNamesLimits) -> Result<Self> {
        validate_user_names_limits(limits)?;
        let expected_path = preflight_user_names(&bytes, limits)?;
        let (package, user_names_path) = open_package(bytes)?;
        // The preflight and the common owner inspect the same immutable input;
        // keep the exact spelling selected by the owner for the replacement.
        if user_names_path != expected_path {
            return Err(Error::PatchConflict {
                operation: "edit user names",
                reason: "User Names stream path changed between CFB validation passes",
            });
        }
        Self::from_package(package, user_names_path, limits)
    }

    fn from_package(
        package: ObjectPackage,
        user_names_path: Vec<String>,
        limits: UserNamesLimits,
    ) -> Result<Self> {
        let bytes = package.source_shared();
        let bytes = Arc::<[u8]>::from(bytes.as_slice());
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
pub struct Edit {
    base: crate::scoped::Base,
    package: Option<crate::Package>,
    source: Option<Snapshot>,
    user_names: Option<UserNamesEdit>,
    options: EditOptions,
}

impl Edit {
    fn lazy(package: crate::Package, options: EditOptions) -> Self {
        Self {
            base: Some(package.clone()),
            package: Some(package),
            source: None,
            user_names: None,
            options,
        }
    }

    fn ensure_open(&mut self) -> Result<()> {
        if self.user_names.is_some() {
            return Ok(());
        }
        let package = self
            .package
            .clone()
            .ok_or_else(|| Error::InvalidData("User Names edit lost its source package".into()))?;
        let (_, bytes) = crate::scoped::capture(&package)?;
        let source = Snapshot::from_bytes(bytes, self.options.limits_value())?;
        self.user_names = Some(source.inner.user_names.edit());
        self.source = Some(source);
        Ok(())
    }

    /// The immutable package snapshot used as the publication source.
    #[must_use]
    pub fn source(&mut self) -> Result<&Snapshot> {
        self.ensure_open()?;
        self.source
            .as_ref()
            .ok_or_else(|| Error::InvalidData("User Names edit source was not retained".into()))
    }

    /// The immutable package snapshot used as the publication source.
    #[must_use]
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.source()
    }

    /// Borrow staged User Names entries in source order.
    #[must_use]
    pub fn users(&mut self) -> Result<&[super::UserEntry]> {
        self.ensure_open()?;
        self.user_names
            .as_ref()
            .map(UserNamesEdit::users)
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))
    }

    /// Borrow the staged User Names semantic projection.
    #[must_use]
    pub fn model(&mut self) -> Result<&super::UserNames> {
        self.ensure_open()?;
        self.user_names
            .as_ref()
            .map(UserNamesEdit::model)
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))
    }

    /// Stage a User Names display-name change.
    ///
    /// # Errors
    ///
    /// Returns an index, name, or bounded stream error without changing this
    /// transaction.
    pub fn set_user_name(&mut self, index: usize, name: impl Into<String>) -> Result<&mut Self> {
        self.ensure_open()?;
        self.user_names
            .as_mut()
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?
            .set_user_name(index, name)?;
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
        self.ensure_open()?;
        self.user_names
            .as_mut()
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?
            .set_user_opened_at(index, opened_at)?;
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
        self.ensure_open()?;
        self.user_names
            .as_mut()
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?
            .insert(index, user)?;
        Ok(self)
    }

    /// Append a User Names entry.
    ///
    /// # Errors
    ///
    /// Returns a GUID-closure, duplicate-ID, or bound error without changing
    /// this transaction.
    pub fn push(&mut self, user: super::UserEntry) -> Result<&mut Self> {
        self.ensure_open()?;
        self.user_names
            .as_mut()
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?
            .push(user)?;
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
        self.ensure_open()?;
        self.user_names
            .as_mut()
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?
            .remove(index)
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
        let base = self.base.clone();
        let (_, patch) = self.commit_record()?;
        let package = crate::scoped::publish(base, patch.before(), patch.after())?;
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(package, patch, diagnostics))
    }

    pub(crate) fn commit_record(mut self) -> Result<(Snapshot, Patch)> {
        self.ensure_open()?;
        let Self {
            source, user_names, ..
        } = self;
        let source = source
            .ok_or_else(|| Error::InvalidData("User Names edit source was not retained".into()))?;
        let user_names = user_names
            .ok_or_else(|| Error::InvalidData("User Names edit was not retained".into()))?;
        let user_names_commit = user_names.commit()?;
        if !user_names_commit.diagnostics().changed() {
            return Ok((
                source.clone(),
                Patch::new(source.bytes_shared(), source.bytes_shared()),
            ));
        }

        let replacement =
            Arc::<[u8]>::from(user_names_commit.package().finish().into_boxed_slice());
        let package = ObjectPackage::open_with_options(
            source.bytes().to_vec(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(PackageLimits::default()),
        )?;
        let mut editor = package.edit()?;
        editor.put_stream_shared(&source.inner.user_names_path, replacement)?;
        let package = editor.commit()?.into_package();
        let snapshot = Snapshot::from_package(
            package,
            source.inner.user_names_path.clone(),
            source.inner.user_names_limits,
        )?;
        let patch = Patch::new(source.bytes_shared(), snapshot.bytes_shared());
        Ok((snapshot, patch))
    }
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("UserNamesPackageEdit")
            .field("materialized", &self.source.is_some())
            .finish_non_exhaustive()
    }
}

/// A published User Names edit: the new package, its exact complete-CFB
/// patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<crate::Package, Patch>;

/// Content-free diagnostics of one User Names commit.
pub use litchi_core::edit::Diagnostics;

/// A User Names edit limit dimension: the bounds of [`super::Limits`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// The `User Names` and Revision Log stream bytes.
    StreamBytes,
    /// The users in the `User Names` stream.
    Users,
    /// The Revision Log GUIDs of the dependency closure.
    RevisionGuids,
}

/// Per-edit configuration of a User Names edit.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: UserNamesLimits,
}

impl EditOptions {
    /// Default options: the default `User Names` bounds.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Select the `User Names` stream bounds.
    #[must_use]
    pub const fn limits(mut self, limits: UserNamesLimits) -> Self {
        self.limits = limits;
        self
    }

    /// The selected `User Names` stream bounds.
    #[must_use]
    pub const fn limits_value(&self) -> UserNamesLimits {
        self.limits
    }
}

impl crate::Package {
    /// The `User Names` read view of this shared workbook.
    ///
    /// # Errors
    ///
    /// Returns an error when the package has no valid `User Names` stream or
    /// its Revision Log dependency closure is invalid.
    pub fn user_names(&self) -> Result<Snapshot> {
        let (_, bytes) = crate::scoped::capture(self)?;
        Snapshot::from_bytes(bytes, UserNamesLimits::default())
    }

    /// Start a `User Names` edit of this package with the default bounds.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::user_names`], or a refusal for an
    /// encrypted workbook.
    pub fn edit_user_names(&self) -> Result<Edit> {
        self.edit_user_names_with_options(EditOptions::default())
    }

    /// Start a `User Names` edit of this package under explicit bounds.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::edit_user_names`], or a limit error
    /// when the package exceeds `options`.
    pub fn edit_user_names_with_options(&self, options: EditOptions) -> Result<Edit> {
        if self.is_encrypted() {
            return Err(Error::refused(
                RefusalKind::Protected,
                "edit user names",
                "an encrypted workbook cannot be republished by a scoped edit",
            ));
        }
        Ok(Edit::lazy(self.clone(), options))
    }

    /// Replay a `User Names` patch onto the exact package it was made from.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// source, or [`Error::Signed`] for a changed signed workbook.
    pub fn apply_user_names(&self, patch: &Patch) -> Result<Commit> {
        crate::scoped::replay_source(self, patch.before(), "edit user names")?;
        let package = crate::scoped::publish(Some(self.clone()), patch.before(), patch.after())?;
        Ok(Commit::new(
            package,
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
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

    /// Apply this patch only to its exact source package snapshot.
    ///
    /// # Errors
    ///
    /// Returns a stale-source error or rejects the target during complete
    /// package reopen.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before() {
            return Err(Error::PatchConflict {
                operation: "edit user names",
                reason: "User Names package patch source does not match its base snapshot",
            });
        }
        if self.is_noop() {
            return Ok(source.clone());
        }
        Snapshot::from_bytes(self.after.to_vec(), source.inner.user_names_limits)
    }

    /// Apply the inverse replacement to its exact committed target.
    ///
    /// # Errors
    ///
    /// Returns a stale-target error or rejects the source during complete
    /// package reopen.
    pub fn revert(&self, target: &Snapshot) -> Result<Snapshot> {
        if target.bytes() != self.after() {
            return Err(Error::PatchConflict {
                operation: "edit user names",
                reason: "User Names package patch target does not match its committed snapshot",
            });
        }
        if self.is_noop() {
            return Ok(target.clone());
        }
        Snapshot::from_bytes(self.before.to_vec(), target.inner.user_names_limits)
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
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())?;
    let path = find_user_names_path(&ole)?;
    let maximum = u64::try_from(limits.max_stream_bytes).unwrap_or(u64::MAX);
    let revision_path = find_revision_log_path(&ole)?;
    let revision_len = ole.stream_len(&[revision_path[0].as_str()])?;
    if revision_len > maximum {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit user names",
            "the Revision Log stream exceeds its byte limit",
        ));
    }
    let stream_len = ole.stream_len(&[path[0].as_str()])?;
    if stream_len > maximum {
        return Err(Error::refused(
            RefusalKind::UnsupportedOperation,
            "edit user names",
            "the User Names stream exceeds its byte limit",
        ));
    }
    Ok(path)
}

fn open_package(bytes: Vec<u8>) -> Result<(ObjectPackage, Vec<String>)> {
    let (package, ole) = ObjectPackage::open_with_ole_file(
        bytes,
        OpenOptions::new()
            .targets(Targets::default())
            .limits(PackageLimits::default()),
    )?;
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
    let mut workbook = crate::raw::Reader::new(Cursor::new(bytes))?;
    workbook
        .user_names(limits)?
        .ok_or_else(|| Error::InvalidData("XLS package has no User Names stream".to_string()))
}
