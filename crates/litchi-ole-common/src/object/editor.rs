//! Transactional CFB stream and selected-storage editing.

use super::cfb_path::CfbPath;
use super::codec::{self, Package as CapturedPackage};
use super::discovery;
use super::link::{self, Link};
use super::model::{self, Limits, Objects};
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::{Package, PackageState};
use super::target::{Target, Targets};
use crate::ole_streams::{self, native, presentation};
use crate::{Error, LimitKind};
use litchi_cfb::{OleError, OleFile, OleFileLimits, SectorLayoutPolicy};
use std::fmt;
use std::io::{Cursor, Read, Seek};
use std::sync::Arc;

/// Maximum number of stream selectors accepted by one removal publication.
pub const MAX_STREAM_REMOVALS: usize = 1_024;

/// A bounded, source- and target-bound replacement prepared for publication.
///
/// The value is intentionally opaque: callers can admit a replacement before
/// cloning an outer format editor, then hand it back to the same object owner
/// for atomic publication. It cannot be constructed without the common
/// owner's target, limits, and validated replacement package.
pub struct PreparedReplacement {
    key: String,
    path: Vec<String>,
    before: Arc<[u8]>,
    replacement: CapturedPackage,
    source: Arc<Vec<u8>>,
    limits: Limits,
}

impl fmt::Debug for PreparedReplacement {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("PreparedReplacement")
            .field("path_depth", &self.path.len())
            .field("has_key", &true)
            .field("has_before", &true)
            .field("has_replacement", &true)
            .field("has_source", &true)
            .finish()
    }
}

/// Result class attached to a completed top-level in-memory CFB parse.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CfbParseOutcome {
    /// The top-level `OleFile::open` call returned a parsed container.
    Success,
    /// The top-level `OleFile::open` call returned an error.
    Error,
}

/// Content-free events emitted around one operation-local top-level in-memory
/// CFB parse.
///
/// The observer is called synchronously and is owned by the caller. The
/// library does not retain it, consult ambient runtime state, or expose source
/// content, offsets, stream names, or physical identifiers.
#[cfg(feature = "performance-diagnostics")]
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CfbParseEvent {
    /// The top-level `OleFile::open` call is about to begin.
    Started,
    /// The top-level `OleFile::open` call has returned.
    Finished { outcome: CfbParseOutcome },
}

#[cfg(feature = "performance-diagnostics")]
fn observe_cfb_parse<T>(
    observer: &mut impl FnMut(CfbParseEvent),
    operation: impl FnOnce() -> Result<T, Error>,
) -> Result<T, Error> {
    observer(CfbParseEvent::Started);
    let result = operation();
    observer(CfbParseEvent::Finished {
        outcome: if result.is_ok() {
            CfbParseOutcome::Success
        } else {
            CfbParseOutcome::Error
        },
    });
    result
}

/// A staged edit of target-selected OLE storages.
///
/// Dropping the edit discards it; the source bytes or snapshot it started
/// from are never changed.
#[derive(Clone)]
pub struct Edit {
    state: Arc<PackageState>,
    /// The captured package the edit started from. This belongs to the
    /// mutable edit, not the immutable owner state: a published `Package`
    /// always uses its current capture as both its base and its package.
    base_package: Arc<CapturedPackage>,
    /// The rendered candidate already validated by the last mutating
    /// operation.  Keeping this publication behind an `Arc` lets the outer
    /// owner hand the same bytes to its source-backed reader when it commits,
    /// without rendering the candidate a second time.
    pending_render: Option<Arc<Vec<u8>>>,
}

impl fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("target_count", &self.state.targets.len())
            .field("object_count", &self.state.objects.len())
            .field("changed", &self.state.changed)
            .field(
                "catalog_is_shared",
                &Arc::ptr_eq(&self.base_package, &self.state.package),
            )
            .finish()
    }
}

impl Edit {
    fn state_mut(&mut self) -> &mut PackageState {
        Arc::make_mut(&mut self.state)
    }

    fn package_mut(&mut self) -> &mut Arc<CapturedPackage> {
        &mut self.state_mut().package
    }

    fn targets_mut(&mut self) -> &mut Targets {
        self.state_mut().targets.make_mut()
    }

    /// Opens a package with an explicit low-level CFB admission profile.
    ///
    /// Security-sensitive format owners should use this entry point when the
    /// source has already been admitted to a caller-owned input ceiling. The
    /// CFB profile is applied before directory, FAT, or selected-object
    /// allocations, then the object capture profile is applied to the admitted
    /// package.
    pub(crate) fn open_with_cfb_limits(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self, Error> {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = OleFile::open(Cursor::new(bytes), cfb_limits)?;
        let admitted = Self::admit_open(&mut ole, targets, limits)?;
        let original = Arc::new(ole.into_inner().into_inner());
        Ok(Self::from_admitted(original, limits, cfb_limits, admitted))
    }

    /// Internal bounded implementation for [`super::Package`] admission.
    pub(crate) fn open_with_ole_file_with_admission_and_cfb_limits<E, F>(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
        cfb_limits: OleFileLimits,
        admission: F,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), E>
    where
        E: From<Error>,
        F: FnOnce(&OleFile<Cursor<Vec<u8>>>) -> Result<(), E>,
    {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = OleFile::open(Cursor::new(bytes), cfb_limits)
            .map_err(Error::from)
            .map_err(E::from)?;
        admission(&ole)?;
        let admitted = Self::admit_open(&mut ole, targets, limits).map_err(E::from)?;
        let original = Arc::new(ole.get_ref().get_ref().clone());
        let editor = Self::from_admitted(original, limits, cfb_limits, admitted);
        Ok((editor, ole))
    }

    #[cfg(feature = "performance-diagnostics")]
    pub(crate) fn open_with_ole_file_profiled_with_admission_and_cfb_limits<E, F>(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
        cfb_limits: OleFileLimits,
        mut observer: impl FnMut(CfbParseEvent),
        admission: F,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), E>
    where
        E: From<Error>,
        F: FnOnce(&OleFile<Cursor<Vec<u8>>>) -> Result<(), E>,
    {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = observe_cfb_parse(&mut observer, || {
            OleFile::open(Cursor::new(bytes), cfb_limits).map_err(Error::from)
        })
        .map_err(E::from)?;
        admission(&ole)?;
        let admitted = Self::admit_open(&mut ole, targets, limits).map_err(E::from)?;
        let original = Arc::new(ole.get_ref().get_ref().clone());
        let editor = Self::from_admitted(original, limits, cfb_limits, admitted);
        Ok((editor, ole))
    }

    fn validate_open_inputs(targets: &Targets, limits: Limits) -> Result<(), Error> {
        limits.validate()?;
        if targets.len() > limits.max_objects {
            return Err(Error::Limit {
                kind: LimitKind::Objects,
                observed: targets.len() as u64,
                maximum: limits.max_objects as u64,
            });
        }
        Ok(())
    }

    fn admit_open<R: Read + Seek>(
        ole: &mut OleFile<R>,
        targets: Targets,
        limits: Limits,
    ) -> Result<(Targets, CapturedPackage, Objects), Error> {
        codec::open(ole, limits.max_package_directory_entries())?;
        if let Some(target) = targets
            .iter()
            .find(|target| target.path().len() > limits.max_storage_depth)
        {
            return Err(Error::Limit {
                kind: LimitKind::StorageDepth,
                observed: target.path().len() as u64,
                maximum: limits.max_storage_depth as u64,
            });
        }

        let resolved_target_entries = targets
            .into_vec()
            .into_iter()
            .map(|target| target.resolve(ole, limits.max_package_directory_entries()))
            .collect::<Result<Vec<_>, _>>()?;
        let resolved_targets = Targets::new(resolved_target_entries)?;
        let package = CapturedPackage::capture(&mut *ole, limits)?;
        package.check(limits)?;
        let objects = discovery::from_package(&package, &resolved_targets, limits)?;
        Ok((resolved_targets, package, objects))
    }

    fn from_admitted(
        original: Arc<Vec<u8>>,
        limits: Limits,
        cfb_limits: OleFileLimits,
        (targets, package, objects): (Targets, CapturedPackage, Objects),
    ) -> Self {
        let targets = targets.into_shared();
        let package = Arc::new(package);
        Self {
            state: Arc::new(PackageState {
                targets,
                limits,
                cfb_limits,
                original: Arc::clone(&original),
                package: Arc::clone(&package),
                objects: objects.into_shared(),
                changed: false,
                layout: SectorLayoutPolicy::default(),
                validated_layout: SectorLayoutPolicy::default(),
            }),
            base_package: package,
            pending_render: None,
        }
    }

    /// A read view of the staged state as an immutable, shareable package.
    ///
    /// The view shares large stream allocations, renders nothing, and is
    /// independent of this edit's subsequent changes; publishing still goes
    /// through [`Self::commit`].
    #[must_use]
    pub(crate) fn staged(&self) -> Package {
        if !self.state.changed {
            return Package::from_state(Arc::clone(&self.state));
        }
        self.staged_with_current(Arc::clone(&self.state.original))
    }

    pub(crate) fn staged_with_current(&self, current: Arc<Vec<u8>>) -> Package {
        Package::new(
            self.state.targets.clone(),
            self.state.limits,
            self.state.cfb_limits,
            // A Package is an immutable publication of the supplied source
            // handle.  Keep that handle separate from the editor's older
            // copy-through base: the published package owns its capture, and
            // a later Package::edit records that capture as its new base.
            Arc::clone(&current),
            // A published package is a fresh immutable base. Do not carry
            // the editor's historical base or its `changed` marker into a
            // later edit: doing so would make an edit started from a prior
            // commit render against stale topology and could turn a no-op
            // follow-up into a byte-changing publication.
            Arc::clone(&self.state.package),
            self.state.objects.clone(),
            false,
            self.state.layout,
            self.state.layout,
        )
    }

    fn into_staged_with_current(mut self, current: Arc<Vec<u8>>) -> Package {
        let state = Arc::make_mut(&mut self.state);
        state.original = Arc::clone(&current);
        state.changed = false;
        state.validated_layout = state.layout;
        Package::from_state(self.state)
    }

    pub(crate) fn from_package(snapshot: &Package) -> Self {
        Self {
            state: snapshot.state_shared(),
            base_package: Arc::clone(&snapshot.state.package),
            pending_render: None,
        }
    }

    pub(crate) fn from_package_with_options(
        package: &Package,
        limits: Limits,
        cfb_limits: OleFileLimits,
        layout: SectorLayoutPolicy,
    ) -> Self {
        let mut edit = Self::from_package(package);
        if edit.state.limits != limits
            || edit.state.cfb_limits != cfb_limits
            || edit.state.layout != layout
        {
            let state = edit.state_mut();
            state.limits = limits;
            state.cfb_limits = cfb_limits;
            state.layout = layout;
        }
        edit
    }

    /// Object capture bounds retained by this editor and its snapshots.
    #[must_use]
    pub fn limits(&self) -> Limits {
        self.state.limits
    }

    /// Physical CFB bounds retained for changed candidate validation.
    #[must_use]
    pub fn cfb_limits(&self) -> OleFileLimits {
        self.state.cfb_limits
    }

    /// The target catalog used by this editor.
    #[must_use]
    pub fn targets(&self) -> &Targets {
        &self.state.targets
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        &self.state.objects
    }

    /// Borrows an opaque package stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.state.package.stream(path)
    }

    /// Returns shared ownership of an opaque package stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.state.package.stream_shared(path)
    }

    /// Returns shared ownership of the exact source bytes admitted by this
    /// editor's bounded open validation.
    ///
    /// Cloning the returned handle does not copy the source allocation. The
    /// bytes remain immutable; edits publish through a separate candidate.
    #[must_use]
    pub fn source_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.state.original)
    }

    /// Applies a fallible replacement to one selected object's standalone CFB.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected target is missing, the callback
    /// rejects the replacement, or the replacement fails CFB validation.
    pub fn update<F>(&mut self, key: &str, edit: F) -> Result<(), Error>
    where
        F: FnOnce(&mut Vec<u8>) -> Result<(), Error>,
    {
        let mut bytes = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .compound()
            .to_vec();
        edit(&mut bytes)?;
        self.replace(key, bytes)
    }

    /// Replaces one selected storage with a validated standalone CFB file.
    ///
    /// # Errors
    ///
    /// Returns an error when the target or replacement is invalid, protected,
    /// oversized, or cannot be committed atomically.
    pub fn replace(&mut self, key: &str, compound_file: Vec<u8>) -> Result<(), Error> {
        let Some(prepared) = self.prepare_replacement(key, compound_file)? else {
            return Ok(());
        };
        self.replace_prepared(prepared)
    }

    /// Validates and captures a replacement without cloning the editor.
    ///
    /// `None` is an exact byte-for-byte no-op. A returned value is bound to
    /// this editor's target, source object, and limits and must be consumed by
    /// [`Self::replace_prepared`] before the editor is changed.
    ///
    /// # Errors
    ///
    /// Returns an error when the target is missing, the replacement exceeds
    /// the retained limits, or the replacement CFB is malformed/protected.
    pub fn prepare_replacement(
        &self,
        key: &str,
        compound_file: Vec<u8>,
    ) -> Result<Option<PreparedReplacement>, Error> {
        if compound_file.len() as u64 > self.state.limits.max_object_size {
            return Err(Error::Limit {
                kind: LimitKind::ObjectBytes,
                observed: compound_file.len() as u64,
                maximum: self.state.limits.max_object_size,
            });
        }
        let object = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?;
        if object.compound() == compound_file.as_slice() {
            return Ok(None);
        }
        let mut replacement_ole = OleFile::open(Cursor::new(compound_file), self.state.cfb_limits)?;
        codec::open(
            &replacement_ole,
            self.state
                .limits
                .max_object_storages()
                .saturating_add(self.state.limits.max_object_streams()),
        )?;
        let replacement = CapturedPackage::capture_object(&mut replacement_ole, self.state.limits)?;
        replacement.check_object_limits(self.state.limits)?;
        Ok(Some(PreparedReplacement {
            key: key.to_owned(),
            path: object.path().to_vec(),
            before: object.compound_shared(),
            replacement,
            source: Arc::clone(&self.state.original),
            limits: self.state.limits,
        }))
    }

    /// Publishes a replacement previously admitted by this editor.
    ///
    /// The prepared value is source-, target-, and limit-bound. Reusing it
    /// with another editor or after the selected object changed is rejected
    /// before any candidate is published.
    ///
    /// # Errors
    ///
    /// Returns an error when the prepared value does not match this editor or
    /// the candidate package fails its normal bounded publication checks.
    pub fn replace_prepared(&mut self, prepared: PreparedReplacement) -> Result<(), Error> {
        if self.state.limits != prepared.limits {
            return Err(OleError::InvalidFormat(
                "prepared replacement limits do not match editor".into(),
            )
            .into());
        }
        if !Arc::ptr_eq(&self.state.original, &prepared.source) {
            return Err(OleError::InvalidFormat(
                "prepared replacement source does not match editor".into(),
            )
            .into());
        }
        let object = self
            .state
            .objects
            .get(&prepared.key)
            .ok_or_else(Error::target_not_found)?;
        if object.path() != prepared.path || object.compound() != prepared.before.as_ref() {
            return Err(OleError::InvalidFormat(
                "prepared replacement source does not match object editor".into(),
            )
            .into());
        }
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).replace_object(
            &prepared.path,
            &prepared.replacement,
            self.state.limits,
        )?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Replaces an existing opaque package stream atomically.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is missing, oversized, or the rendered
    /// package fails validation.
    pub fn put_stream(&mut self, path: &[String], data: Vec<u8>) -> Result<(), Error> {
        self.put_stream_shared(path, data.into())
    }

    /// Replaces an existing stream while retaining the caller's allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is missing, oversized, or the rendered
    /// package fails validation.
    pub fn put_stream_shared(&mut self, path: &[String], data: Arc<[u8]>) -> Result<(), Error> {
        if self
            .stream(path)
            .is_some_and(|current| current == data.as_ref())
        {
            return Ok(());
        }
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).put_stream(path, data, self.state.limits)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Replaces existing streams in one failure-atomic CFB publication.
    ///
    /// Replacements are applied in iterator order to one isolated candidate.
    /// Repeated paths therefore retain the last supplied value. The candidate
    /// is rendered, reopened, and published only once after every replacement
    /// has passed the configured stream and package bounds.
    ///
    /// # Errors
    ///
    /// Returns an error when the batch exceeds the package stream-count bound,
    /// a stream is missing or oversized, or the final rendered package fails
    /// validation. Failure leaves this editor unchanged.
    pub fn put_streams_shared<'a>(
        &mut self,
        replacements: impl IntoIterator<Item = (&'a [String], Arc<[u8]>)>,
    ) -> Result<(), Error> {
        let mut candidate = None;
        for (index, (path, data)) in replacements.into_iter().enumerate() {
            if index >= self.state.limits.max_streams {
                return Err(Error::Limit {
                    kind: LimitKind::PackageStreams,
                    observed: index.saturating_add(1) as u64,
                    maximum: self.state.limits.max_streams as u64,
                });
            }
            if candidate
                .as_ref()
                .map_or_else(|| self.stream(path), |editor: &Self| editor.stream(path))
                .is_some_and(|current| current == data.as_ref())
            {
                continue;
            }
            let candidate = candidate.get_or_insert_with(|| self.clone());
            Arc::make_mut(candidate.package_mut()).put_stream(path, data, self.state.limits)?;
        }
        if let Some(candidate) = candidate {
            *self = candidate.commit_candidate()?;
        }
        Ok(())
    }

    /// Adds an opaque package stream below an existing CFB storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the path or data exceeds the configured bounds,
    /// the parent is missing, or the rendered package fails validation.
    pub fn add_stream(&mut self, path: Vec<String>, data: Vec<u8>) -> Result<(), Error> {
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).add_stream(path, data.into(), self.state.limits)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Removes one opaque package stream while retaining its parent storage.
    ///
    /// Paths use CFB's Unicode simple-uppercase identity and are resolved to
    /// the spelling stored in the package. An absent stream is an exact no-op
    /// represented by `Ok(None)`; an existing storage at the same path is not
    /// treated as absent.
    ///
    /// # Errors
    ///
    /// Returns an error when the path is empty or invalid, identifies a
    /// storage, or the rendered package fails validation. Failure leaves this
    /// editor unchanged.
    pub fn remove_stream(&mut self, path: &[String]) -> Result<Option<Arc<[u8]>>, Error> {
        self.validate_stream_path_depth(path)?;
        let path = CfbPath::try_from_slice(path, "stream removal path")
            .map_err(|_| Error::invalid_selector())?;
        let Some(removed) = self.state.package.removable_stream(&path)? else {
            return Ok(None);
        };
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).remove_stream(&path, self.state.limits)?;
        *self = candidate.commit_candidate()?;
        Ok(Some(removed))
    }

    /// Removes multiple opaque package streams in one failure-atomic publish.
    ///
    /// Results correspond positionally to the supplied selectors. Missing
    /// streams yield `None`, while present streams yield their shared bytes.
    /// CFB-equivalent duplicate selectors are refused instead of depending on
    /// iterator order. Empty batches and all-absent batches are exact no-ops.
    ///
    /// # Errors
    ///
    /// Returns an error when the batch exceeds [`MAX_STREAM_REMOVALS`], a path
    /// is invalid or duplicated, a selector identifies a storage, allocation
    /// fails, or the rendered package fails validation. Failure leaves this
    /// editor unchanged.
    pub fn remove_streams<'a>(
        &mut self,
        paths: impl IntoIterator<Item = &'a [String]>,
    ) -> Result<Vec<Option<Arc<[u8]>>>, Error> {
        let mut validated = Vec::<(CfbPath, u64)>::new();
        for path in paths {
            if validated.len() == MAX_STREAM_REMOVALS {
                return Err(Error::Limit {
                    kind: LimitKind::StreamRemovals,
                    observed: validated.len().saturating_add(1) as u64,
                    maximum: MAX_STREAM_REMOVALS as u64,
                });
            }
            self.validate_stream_path_depth(path)?;
            let path = CfbPath::try_from_slice(path, "stream removal path")
                .map_err(|_| Error::invalid_selector())?;
            let identity = path.identity_hash();
            if validated.iter().any(|(existing, existing_identity)| {
                *existing_identity == identity && existing.same_as(&path)
            }) {
                return Err(Error::invalid_selector());
            }
            validated
                .try_reserve(1)
                .map_err(|source| OleError::Allocation {
                    resource: "stream removal selectors",
                    source,
                })?;
            validated.push((path, identity));
        }
        if validated.is_empty() {
            return Ok(Vec::new());
        }
        let mut candidate = self.clone();
        let removed = Arc::make_mut(candidate.package_mut()).remove_streams(
            validated.iter().map(|(path, _identity)| path),
            self.state.limits,
        )?;
        let changed = removed.iter().any(Option::is_some);
        if changed {
            *self = candidate.commit_candidate()?;
        }
        Ok(removed)
    }

    fn validate_stream_path_depth(&self, path: &[String]) -> Result<(), Error> {
        let maximum = self.state.limits.max_storage_depth.saturating_add(1);
        if path.len() > maximum {
            return Err(Error::Limit {
                kind: LimitKind::SelectorDepth,
                observed: path.len() as u64,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    /// Atomically edits one selected object's OLEDS `\x01Ole` metadata.
    ///
    /// The callback only receives the inert typed link fields.  The edited
    /// stream is published through the same candidate-render-and-reopen path
    /// as every other object edit, so callback or CFB failures leave this
    /// editor unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when `key` is absent, its OLEDS stream is missing or
    /// malformed, `edit` fails, or the resulting package cannot be validated.
    pub fn update_link<F>(&mut self, key: &str, limits: link::Limits, edit: F) -> Result<(), Error>
    where
        F: FnOnce(&mut Link) -> Result<(), Error>,
    {
        let object_path = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?
            .path()
            .to_vec();
        let mut stream_path = object_path;
        stream_path.push(link::NAME.to_string());
        let bytes = self
            .state
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        let mut link = Link::parse_shared(bytes, limits)?;
        edit(&mut link)?;
        self.put_stream(&stream_path, link.to_bytes(limits)?)
    }

    /// Atomically edits one selected object's OLEDS presentation under
    /// explicit stream limits.
    pub fn update_presentation<F>(
        &mut self,
        key: &str,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), Error>
    where
        F: FnOnce(&mut presentation::Edit) -> Result<(), Error>,
    {
        model::validate_ole_stream_limits(limits)?;
        let stream_path = self.object_presentation_path(key, index)?;
        let bytes = self
            .state
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        let source = presentation::Snapshot::parse_shared(bytes, limits)?;
        let mut transaction = source.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        self.put_stream_shared(&stream_path, commit.package().bytes_shared())
    }

    /// Applies a source-checked OLEDS presentation patch to one selected
    /// object atomically.
    ///
    /// The exact source bytes are compared before reparsing, so a stale patch
    /// is rejected before an editor candidate is cloned or rendered.
    pub fn apply_presentation_patch(
        &mut self,
        key: &str,
        index: usize,
        patch: &presentation::Patch,
    ) -> Result<(), Error> {
        let stream_path = self.object_presentation_path(key, index)?;
        let bytes = self
            .state
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        if bytes.as_ref() != patch.before_bytes() {
            return Err(Error::PatchConflict);
        }
        let source = presentation::Snapshot::parse_shared(bytes, patch.source().limits())?;
        let replacement = patch.apply(&source)?;
        self.put_stream_shared(&stream_path, replacement.bytes_shared())
    }

    /// Atomically edits one selected object's native-data stream under
    /// explicit stream limits.
    pub fn update_native<F>(
        &mut self,
        key: &str,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), Error>
    where
        F: FnOnce(&mut native::Edit) -> Result<(), Error>,
    {
        model::validate_ole_stream_limits(limits)?;
        let stream_path = self.object_native_path(key)?;
        let bytes = self
            .state
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        let source = native::Snapshot::parse_shared(bytes, limits)?;
        let mut transaction = source.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        self.put_stream_shared(&stream_path, commit.package().bytes_shared())
    }

    /// Applies a source-checked OLEDS native-data patch to one selected object
    /// atomically.
    ///
    /// The exact source bytes are compared before reparsing, so a stale patch
    /// is rejected before an editor candidate is cloned or rendered.
    pub fn apply_native_patch(&mut self, key: &str, patch: &native::Patch) -> Result<(), Error> {
        let stream_path = self.object_native_path(key)?;
        let bytes = self
            .state
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        if bytes.as_ref() != patch.before_bytes() {
            return Err(Error::PatchConflict);
        }
        let source = native::Snapshot::parse_shared(bytes, patch.source().limits())?;
        let replacement = patch.apply(&source)?;
        self.put_stream_shared(&stream_path, replacement.bytes_shared())
    }

    fn object_presentation_path(&self, key: &str, index: usize) -> Result<Vec<String>, Error> {
        let object = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?;
        object.validate_presentation_count()?;
        let mut path = object.path().to_vec();
        path.push(ole_streams::presentation_name(index)?);
        Ok(path)
    }

    fn object_native_path(&self, key: &str) -> Result<Vec<String>, Error> {
        let object = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?;
        object.validate_presentation_count()?;
        let mut path = object.path().to_vec();
        path.push(ole_streams::NATIVE_STREAM_NAME.to_string());
        Ok(path)
    }

    /// Adds a target-selected storage after the host has staged its reference.
    ///
    /// # Errors
    ///
    /// Returns an error when the target or replacement CFB is invalid, the
    /// target already exists, or the rendered package fails validation.
    pub fn add_storage(&mut self, target: Target, compound_file: Vec<u8>) -> Result<(), Error> {
        if compound_file.len() as u64 > self.state.limits.max_object_size {
            return Err(Error::Limit {
                kind: LimitKind::ObjectBytes,
                observed: compound_file.len() as u64,
                maximum: self.state.limits.max_object_size,
            });
        }
        let mut nested_ole = OleFile::open(Cursor::new(compound_file), self.state.cfb_limits)?;
        codec::open(
            &nested_ole,
            self.state
                .limits
                .max_object_storages()
                .saturating_add(self.state.limits.max_object_streams()),
        )?;
        let nested = CapturedPackage::capture_object(&mut nested_ole, self.state.limits)?;
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).add_object(&target, &nested, self.state.limits)?;
        candidate.targets_mut().push(target)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Removes a selected storage after the host has removed its references.
    ///
    /// # Errors
    ///
    /// Returns an error when the target is missing or the rendered package
    /// fails validation.
    pub fn remove_storage(&mut self, key: &str) -> Result<Arc<[u8]>, Error> {
        let object = self
            .state
            .objects
            .get(key)
            .ok_or_else(Error::target_not_found)?;
        let removed = object.compound_shared();
        let mut candidate = self.clone();
        Arc::make_mut(candidate.package_mut()).remove_object(object.path(), self.state.limits)?;
        let targets = candidate.state.targets.without(key)?;
        candidate.state_mut().targets = targets.into_shared();
        *self = candidate.commit_candidate()?;
        Ok(removed)
    }

    /// Selects where a rendered package places its sectors.
    ///
    /// The default is [`SectorLayoutPolicy::Reuse`]: a render keeps the opened
    /// artifact's sector assignment wherever the package's stream and storage
    /// set still matches it, appends what no longer fits, and reclaims what a
    /// shrinking stream released. [`SectorLayoutPolicy::Rewrite`] re-lays out
    /// the whole container, producing the smallest output.
    ///
    /// Both policies publish the same logical package: identical stream bytes,
    /// hierarchy, names and class identifiers.
    pub fn set_sector_layout_policy(&mut self, policy: SectorLayoutPolicy) {
        self.state_mut().layout = policy;
    }

    /// The sector-layout policy a render will apply.
    #[must_use]
    pub fn sector_layout_policy(&self) -> SectorLayoutPolicy {
        self.state.layout
    }

    /// Finishes the edit, returning the original bytes for a true no-op.
    ///
    /// # Errors
    ///
    /// Returns an error if the edited CFB cannot be rendered.
    fn finish(&self) -> Result<Vec<u8>, Error> {
        if self.state.changed {
            if self.state.layout == SectorLayoutPolicy::Reuse
                && let Some(rendered) = self.state.package.render_copy_through(
                    &self.base_package,
                    &self.state.original,
                    self.state.limits,
                    self.state.cfb_limits,
                )?
            {
                return Ok(rendered);
            }
            let rendered = self.state.package.render_with_layout(
                Some(self.state.original.as_slice()),
                self.state.layout,
                self.state.cfb_limits,
            )?;
            // Ordinary publication repeats the already validated deterministic
            // render. A caller-selected layout change needs a new physical gate.
            if self.state.layout != self.state.validated_layout {
                open_candidate_with_limits(rendered.as_slice(), self.state.cfb_limits)?;
            }
            Ok(rendered)
        } else {
            Ok(self.state.original.as_ref().clone())
        }
    }

    /// Commits this edit as an immutable package plus a reversible patch.
    ///
    /// The source editor is consumed, so callers cannot accidentally keep
    /// mutating a value after using the commit result. The patch is checked
    /// against the exact original artifact and the package has already
    /// passed the common CFB/resource validation performed by each edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited package cannot be rendered.
    pub fn commit(self) -> Result<Commit, Error> {
        let before = Arc::clone(&self.state.original);
        let after = if self.state.changed {
            if self.state.layout == self.state.validated_layout
                && let Some(rendered) = self.pending_render.clone()
            {
                rendered
            } else {
                Arc::new(self.finish()?)
            }
        } else {
            Arc::clone(&self.state.original)
        };
        let package = if self.state.changed {
            self.into_staged_with_current(Arc::clone(&after))
        } else {
            Package::from_state(Arc::clone(&self.state))
        };
        let patch = Patch::new_shared(before, after);
        let diagnostics = Diagnostics::new(!patch.is_noop());
        Ok(Commit::new(package, patch, diagnostics))
    }

    fn commit_candidate(mut self) -> Result<Self, Error> {
        self.state.package.check(self.state.limits)?;
        let rendered = if self.state.layout == SectorLayoutPolicy::Reuse {
            self.state
                .package
                .render_copy_through(
                    &self.base_package,
                    &self.state.original,
                    self.state.limits,
                    self.state.cfb_limits,
                )?
                .map_or_else(
                    || {
                        self.state.package.render_with_layout(
                            Some(self.state.original.as_slice()),
                            self.state.layout,
                            self.state.cfb_limits,
                        )
                    },
                    Ok,
                )?
        } else {
            self.state.package.render_with_layout(
                Some(self.state.original.as_slice()),
                self.state.layout,
                self.state.cfb_limits,
            )?
        };
        let mut check = open_candidate_with_limits(rendered.as_slice(), self.state.cfb_limits)?;
        codec::open(&check, self.state.limits.max_package_directory_entries())?;
        let mut parsed = CapturedPackage::capture(&mut check, self.state.limits)?;
        parsed.reuse_stream_allocations(&self.state.package)?;
        let objects = discovery::from_package(&parsed, &self.state.targets, self.state.limits)?;
        let state = self.state_mut();
        // `package_mut` is called before every candidate publication.  Once
        // it has detached the candidate package, its allocation is private
        // to this edit; the base package still owns the prior capture.  Move
        // the validated capture into that allocation instead of allocating a
        // second `Arc<CapturedPackage>`.  Keep the fallback for callers that
        // construct a candidate without first mutating its package.
        if let Some(package) = Arc::get_mut(&mut state.package) {
            *package = parsed;
        } else {
            state.package = Arc::new(parsed);
        }
        state.objects = objects.into_shared();
        state.changed = true;
        state.validated_layout = state.layout;
        self.pending_render = Some(Arc::new(rendered));
        Ok(self)
    }
}

/// Reopen a rendered mutation candidate under the editor's retained physical
/// CFB policy. The common object edit surface reports a candidate that grows
/// beyond that policy as its existing stream-byte budget, while direct CFB
/// ingress keeps the low-level physical label and error intact.
fn open_candidate_with_limits(
    bytes: &[u8],
    limits: OleFileLimits,
) -> Result<OleFile<Cursor<&[u8]>>, Error> {
    match OleFile::open(Cursor::new(bytes), limits) {
        Ok(file) => Ok(file),
        Err(OleError::LimitExceeded {
            resource: "input bytes",
            observed,
            maximum,
        }) => Err(Error::Limit {
            kind: LimitKind::StreamBytes,
            observed,
            maximum,
        }),
        Err(error) => Err(error.into()),
    }
}
