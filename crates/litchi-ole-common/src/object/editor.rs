//! Transactional CFB stream and selected-storage editing.

use super::cfb_path::CfbPath;
use super::codec::{self, Package};
use super::discovery;
use super::link::{self, Link};
use super::model::{self, Limits, Objects};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::target::{Target, Targets};
use crate::ole_streams::{
    self, NativePatch, NativeSnapshot, NativeTransaction, PresentationPatch, PresentationSnapshot,
    PresentationTransaction,
};
use litchi_cfb::{OleError, OleFile, OleFileLimits, SectorLayoutPolicy};
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
#[derive(Debug)]
pub struct PreparedReplacement {
    key: String,
    path: Vec<String>,
    before: Arc<[u8]>,
    replacement: Package,
    source: Arc<Vec<u8>>,
    limits: Limits,
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
    operation: impl FnOnce() -> Result<T, OleError>,
) -> Result<T, OleError> {
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

/// Transactional editor for target-selected OLE storages.
#[derive(Debug, Clone)]
pub struct Editor {
    targets: Targets,
    limits: Limits,
    original: Arc<Vec<u8>>,
    base_package: Package,
    package: Package,
    objects: Objects,
    changed: bool,
    layout: SectorLayoutPolicy,
}

impl Editor {
    /// Opens a CFB package with an explicit host-resolved target catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the CFB is malformed or protected, a target is
    /// missing/invalid, or a configured resource limit is exceeded.
    pub fn open(bytes: Vec<u8>, targets: Targets, limits: Limits) -> Result<Self, OleError> {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = OleFile::open(Cursor::new(bytes))?;
        let admitted = Self::admit_open(&mut ole, targets, limits)?;
        let original = Arc::new(ole.into_inner().into_inner());
        Ok(Self::from_admitted(original, limits, admitted))
    }

    /// Opens a package with an explicit low-level CFB admission profile.
    ///
    /// Security-sensitive format owners should use this entry point when the
    /// source has already been admitted to a caller-owned input ceiling. The
    /// CFB profile is applied before directory, FAT, or selected-object
    /// allocations, then the object capture profile is applied to the admitted
    /// package.
    pub fn open_with_cfb_limits(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
        cfb_limits: OleFileLimits,
    ) -> Result<Self, OleError> {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = OleFile::open_with_limits(Cursor::new(bytes), cfb_limits)?;
        let admitted = Self::admit_open(&mut ole, targets, limits)?;
        let original = Arc::new(ole.into_inner().into_inner());
        Ok(Self::from_admitted(original, limits, admitted))
    }

    /// Opens a package and returns the validated physical OLE context used by
    /// the object owner.
    ///
    /// The returned `OleFile` is parsed from the same source content as the
    /// returned editor. Keeping source opening inside this bounded API
    /// prevents callers from pairing an editor with an unrelated parsed CFB.
    /// The editor retains one source copy because both values are returned;
    /// that copy is made only after common CFB admission succeeds.
    ///
    /// # Errors
    ///
    /// Returns the same protection, target, capture, discovery, and resource
    /// limit errors as [`Self::open`].
    pub fn open_with_ole_file(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), OleError> {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = OleFile::open(Cursor::new(bytes))?;
        let admitted = Self::admit_open(&mut ole, targets, limits)?;
        let original = Arc::new(ole.get_ref().get_ref().clone());
        let editor = Self::from_admitted(original, limits, admitted);
        Ok((editor, ole))
    }

    /// Opens a package and reports the operation-local top-level in-memory CFB
    /// parse.
    ///
    /// This is the feature-gated equivalent of [`Self::open_with_ole_file`].
    /// When strict-owner input validation permits parsing, the observer
    /// receives exactly one balanced pair around the top-level
    /// `OleFile::open`. Semantic target/package failures occur after a
    /// successful parse and are not folded into the parse result.
    ///
    /// # Panics
    ///
    /// Observer panics propagate. Event balancing is guaranteed only when the
    /// observer returns normally from every callback.
    ///
    /// # Errors
    ///
    /// Returns the same errors as [`Self::open_with_ole_file`].
    #[cfg(feature = "performance-diagnostics")]
    pub fn open_with_ole_file_profiled(
        bytes: Vec<u8>,
        targets: Targets,
        limits: Limits,
        mut observer: impl FnMut(CfbParseEvent),
    ) -> Result<(Self, OleFile<Cursor<Vec<u8>>>), OleError> {
        Self::validate_open_inputs(&targets, limits)?;
        let mut ole = observe_cfb_parse(&mut observer, || OleFile::open(Cursor::new(bytes)))?;
        let admitted = Self::admit_open(&mut ole, targets, limits)?;
        let original = Arc::new(ole.get_ref().get_ref().clone());
        let editor = Self::from_admitted(original, limits, admitted);
        Ok((editor, ole))
    }

    fn validate_open_inputs(targets: &Targets, limits: Limits) -> Result<(), OleError> {
        limits.validate()?;
        if targets.len() > limits.max_objects {
            return Err(OleError::InvalidFormat(format!(
                "object target count {} exceeds limit {}",
                targets.len(),
                limits.max_objects
            )));
        }
        Ok(())
    }

    fn admit_open<R: Read + Seek>(
        ole: &mut OleFile<R>,
        targets: Targets,
        limits: Limits,
    ) -> Result<(Targets, Package, Objects), OleError> {
        codec::open(ole, limits.max_package_directory_entries())?;
        if targets
            .iter()
            .any(|target| target.path().len() > limits.max_storage_depth)
        {
            return Err(OleError::InvalidFormat(
                "object target path exceeds storage depth limit".into(),
            ));
        }
        let resolved_target_entries = targets
            .into_vec()
            .into_iter()
            .map(|target| target.resolve(ole, limits.max_package_directory_entries()))
            .collect::<Result<Vec<_>, _>>()?;
        let resolved_targets = Targets::new(resolved_target_entries)?;
        let package = Package::capture(&mut *ole, limits)?;
        package.check(limits)?;
        let objects = discovery::from_package(&package, &resolved_targets, limits)?;
        Ok((resolved_targets, package, objects))
    }

    fn from_admitted(
        original: Arc<Vec<u8>>,
        limits: Limits,
        (targets, package, objects): (Targets, Package, Objects),
    ) -> Self {
        Self {
            targets,
            limits,
            original,
            base_package: package.clone(),
            package,
            objects,
            changed: false,
            layout: SectorLayoutPolicy::default(),
        }
    }

    /// Captures the current read state as an immutable, shareable snapshot.
    ///
    /// Snapshot clones share large stream allocations and are independent of
    /// this editor's subsequent edits.
    #[must_use]
    pub fn snapshot(&self) -> Snapshot {
        Snapshot::new(
            self.targets.clone(),
            self.limits,
            Arc::clone(&self.original),
            self.base_package.clone(),
            self.package.clone(),
            self.objects.clone(),
            self.changed,
            self.layout,
        )
    }

    pub(crate) fn from_snapshot(snapshot: &Snapshot) -> Self {
        Self {
            targets: snapshot.targets().clone(),
            limits: snapshot.limits(),
            original: snapshot.original(),
            base_package: snapshot.base_package(),
            package: snapshot.package(),
            objects: snapshot.objects_clone(),
            changed: snapshot.changed(),
            layout: snapshot.layout(),
        }
    }

    /// The target catalog used by this editor.
    #[must_use]
    pub fn targets(&self) -> &Targets {
        &self.targets
    }

    /// The current target-selected object catalog.
    #[must_use]
    pub fn objects(&self) -> &Objects {
        &self.objects
    }

    /// Whether a committed edit has changed the package.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrows an opaque package stream without copying it.
    #[must_use]
    pub fn stream(&self, path: &[String]) -> Option<&[u8]> {
        self.package.stream(path)
    }

    /// Returns shared ownership of an opaque package stream allocation.
    #[must_use]
    pub fn stream_shared(&self, path: &[String]) -> Option<Arc<[u8]>> {
        self.package.stream_shared(path)
    }

    /// Returns shared ownership of the exact source bytes admitted by this
    /// editor's bounded open validation.
    ///
    /// Cloning the returned handle does not copy the source allocation. The
    /// bytes remain immutable; edits publish through a separate candidate.
    #[must_use]
    pub fn source_shared(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.original)
    }

    /// Applies a fallible replacement to one selected object's standalone CFB.
    ///
    /// # Errors
    ///
    /// Returns an error when the selected target is missing, the callback
    /// rejects the replacement, or the replacement fails CFB validation.
    pub fn update<F>(&mut self, key: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Vec<u8>) -> Result<(), OleError>,
    {
        let mut bytes = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
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
    pub fn replace(&mut self, key: &str, compound_file: Vec<u8>) -> Result<(), OleError> {
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
    ) -> Result<Option<PreparedReplacement>, OleError> {
        if compound_file.len() as u64 > self.limits.max_object_size {
            return Err(OleError::InvalidFormat(
                "replacement object exceeds size limit".into(),
            ));
        }
        let object = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?;
        if object.compound() == compound_file.as_slice() {
            return Ok(None);
        }
        let mut replacement_ole = OleFile::open(Cursor::new(compound_file))?;
        codec::open(
            &replacement_ole,
            self.limits
                .max_object_storages()
                .saturating_add(self.limits.max_object_streams()),
        )?;
        let replacement = Package::capture_object(&mut replacement_ole, self.limits)?;
        replacement.check_object_limits(self.limits)?;
        Ok(Some(PreparedReplacement {
            key: key.to_owned(),
            path: object.path().to_vec(),
            before: object.compound_shared(),
            replacement,
            source: Arc::clone(&self.original),
            limits: self.limits,
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
    pub fn replace_prepared(&mut self, prepared: PreparedReplacement) -> Result<(), OleError> {
        if self.limits != prepared.limits {
            return Err(OleError::InvalidFormat(
                "prepared replacement limits do not match editor".into(),
            ));
        }
        if !Arc::ptr_eq(&self.original, &prepared.source) {
            return Err(OleError::InvalidFormat(
                "prepared replacement source does not match editor".into(),
            ));
        }
        let object = self.objects.get(&prepared.key).ok_or_else(|| {
            OleError::InvalidFormat(format!("object target {:?} not found", prepared.key))
        })?;
        if object.path() != prepared.path || object.compound() != prepared.before.as_ref() {
            return Err(OleError::InvalidFormat(
                "prepared replacement source does not match object editor".into(),
            ));
        }
        let mut candidate = self.clone();
        candidate
            .package
            .replace_object(&prepared.path, &prepared.replacement, self.limits)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Replaces an existing opaque package stream atomically.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is missing, oversized, or the rendered
    /// package fails validation.
    pub fn put_stream(&mut self, path: &[String], data: Vec<u8>) -> Result<(), OleError> {
        self.put_stream_shared(path, data.into())
    }

    /// Replaces an existing stream while retaining the caller's allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is missing, oversized, or the rendered
    /// package fails validation.
    pub fn put_stream_shared(&mut self, path: &[String], data: Arc<[u8]>) -> Result<(), OleError> {
        if self
            .stream(path)
            .is_some_and(|current| current == data.as_ref())
        {
            return Ok(());
        }
        let mut candidate = self.clone();
        candidate.package.put_stream(path, data, self.limits)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Replaces an existing stream and returns the exact candidate rendering
    /// that was validated while committing it.
    ///
    /// This is a narrow hand-off for format owners whose next validation stage
    /// consumes the rendered package bytes.  It does not cache a rendering on
    /// the editor, so ordinary `put_stream_shared` callers and no-op editors
    /// retain their existing behavior and allocation profile.  The returned
    /// bytes have already passed the same package check, CFB reopen, stream
    /// recapture, and target discovery performed by a normal stream edit.
    ///
    /// # Errors
    ///
    /// Returns the same bounded stream, CFB, and package-validation errors as
    /// [`Self::put_stream_shared`].
    pub fn put_stream_shared_with_rendered(
        &mut self,
        path: &[String],
        data: Arc<[u8]>,
    ) -> Result<Vec<u8>, OleError> {
        if self
            .stream(path)
            .is_some_and(|current| current == data.as_ref())
        {
            return self.clone().finish();
        }
        let mut candidate = self.clone();
        candidate.package.put_stream(path, data, self.limits)?;
        let (candidate, rendered) = candidate.commit_candidate_with_rendered()?;
        *self = candidate;
        Ok(rendered)
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
    ) -> Result<(), OleError> {
        self.put_streams_shared_with_rendered(replacements)
            .map(|_| ())
    }

    /// Replaces existing streams in one failure-atomic CFB publication and
    /// returns the validated candidate rendering when the batch changes it.
    ///
    /// Replacements are applied in iterator order to one isolated candidate.
    /// Repeated paths therefore retain the last supplied value. The candidate
    /// is rendered, reopened, and published only once after every replacement
    /// has passed the configured stream and package bounds. An empty or
    /// all-no-op batch returns `Ok(None)` without cloning or rendering. A
    /// repeated path that changes to one value and then changes back to its
    /// original value is still an effective two-step batch and is rendered.
    ///
    /// The returned bytes are the exact rendering validated while publishing
    /// the candidate. They are handed to the caller and are not retained by
    /// this editor.
    ///
    /// # Errors
    ///
    /// Returns an error when the batch exceeds the package stream-count bound,
    /// a stream is missing or oversized, or the final rendered package fails
    /// validation. Failure leaves this editor unchanged and returns no
    /// rendered bytes.
    pub fn put_streams_shared_with_rendered<'a>(
        &mut self,
        replacements: impl IntoIterator<Item = (&'a [String], Arc<[u8]>)>,
    ) -> Result<Option<Vec<u8>>, OleError> {
        let mut candidate = None;
        for (index, (path, data)) in replacements.into_iter().enumerate() {
            if index >= self.limits.max_streams {
                return Err(OleError::InvalidFormat(
                    "stream replacement batch exceeds package stream-count limit".into(),
                ));
            }
            if candidate
                .as_ref()
                .map_or_else(|| self.stream(path), |editor: &Self| editor.stream(path))
                .is_some_and(|current| current == data.as_ref())
            {
                continue;
            }
            let candidate = candidate.get_or_insert_with(|| self.clone());
            candidate.package.put_stream(path, data, self.limits)?;
        }
        if let Some(candidate) = candidate {
            let (candidate, rendered) = candidate.commit_candidate_with_rendered()?;
            *self = candidate;
            Ok(Some(rendered))
        } else {
            Ok(None)
        }
    }

    /// Adds an opaque package stream below an existing CFB storage.
    ///
    /// # Errors
    ///
    /// Returns an error when the path or data exceeds the configured bounds,
    /// the parent is missing, or the rendered package fails validation.
    pub fn add_stream(&mut self, path: Vec<String>, data: Vec<u8>) -> Result<(), OleError> {
        let mut candidate = self.clone();
        candidate
            .package
            .add_stream(path, data.into(), self.limits)?;
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
    pub fn remove_stream(&mut self, path: &[String]) -> Result<Option<Arc<[u8]>>, OleError> {
        self.validate_stream_path_depth(path)?;
        let path = CfbPath::try_from_slice(path, "stream removal path")?;
        let Some(removed) = self.package.removable_stream(&path)? else {
            return Ok(None);
        };
        let mut candidate = self.clone();
        candidate.package.remove_stream(&path, self.limits)?;
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
    ) -> Result<Vec<Option<Arc<[u8]>>>, OleError> {
        let mut validated = Vec::<(CfbPath, u64)>::new();
        for path in paths {
            if validated.len() == MAX_STREAM_REMOVALS {
                return Err(OleError::InvalidFormat(format!(
                    "stream removal batch exceeds operation limit {MAX_STREAM_REMOVALS}"
                )));
            }
            self.validate_stream_path_depth(path)?;
            let path = CfbPath::try_from_slice(path, "stream removal path")?;
            let identity = path.identity_hash();
            if validated.iter().any(|(existing, existing_identity)| {
                *existing_identity == identity && existing.same_as(&path)
            }) {
                return Err(OleError::InvalidFormat(format!(
                    "stream removal batch contains duplicate path {:?}",
                    path.as_slice()
                )));
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
        let removed = candidate
            .package
            .remove_streams(validated.iter().map(|(path, _identity)| path), self.limits)?;
        let changed = removed.iter().any(Option::is_some);
        if changed {
            *self = candidate.commit_candidate()?;
        }
        Ok(removed)
    }

    fn validate_stream_path_depth(&self, path: &[String]) -> Result<(), OleError> {
        let maximum = self
            .limits
            .max_storage_depth
            .checked_add(1)
            .ok_or_else(|| {
                OleError::InvalidFormat("stream selector depth limit overflows usize".into())
            })?;
        if path.len() > maximum {
            return Err(OleError::InvalidFormat(format!(
                "stream selector depth {} exceeds limit {maximum}",
                path.len()
            )));
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
    pub fn update_link<F>(&mut self, key: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut Link) -> Result<(), OleError>,
    {
        let object_path = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?
            .path()
            .to_vec();
        let mut stream_path = object_path;
        stream_path.push(link::NAME.to_string());
        let bytes = self
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        let mut link = Link::parse_shared(bytes)?;
        edit(&mut link)?;
        self.put_stream(&stream_path, link.to_bytes())
    }

    /// Atomically edits one selected object's OLEDS presentation stream.
    ///
    /// The source stream is parsed into a bounded, shared snapshot before the
    /// callback runs.  The codec transaction validates and reparses the
    /// candidate, then this editor publishes it through the ordinary
    /// clone-render-reopen path.  A callback failure or package validation
    /// failure leaves this editor unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error when the target or indexed stream is absent, the
    /// stream is malformed, the callback rejects the edit, or the resulting
    /// package cannot be validated under the editor limits.
    pub fn update_presentation<F>(
        &mut self,
        key: &str,
        index: usize,
        edit: F,
    ) -> Result<(), OleError>
    where
        F: FnOnce(&mut PresentationTransaction) -> Result<(), OleError>,
    {
        self.update_presentation_with_limits(key, index, ole_streams::Limits::default(), edit)
    }

    /// Atomically edits one selected object's OLEDS presentation under
    /// explicit stream limits.
    pub fn update_presentation_with_limits<F>(
        &mut self,
        key: &str,
        index: usize,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), OleError>
    where
        F: FnOnce(&mut PresentationTransaction) -> Result<(), OleError>,
    {
        model::validate_ole_stream_limits(limits)?;
        let stream_path = self.object_presentation_path(key, index)?;
        let source = self
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)
            .and_then(|bytes| PresentationSnapshot::parse_shared(bytes, limits))?;
        let mut transaction = source.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        self.put_stream_shared(&stream_path, commit.snapshot().bytes_shared())
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
        patch: &PresentationPatch,
    ) -> Result<(), OleError> {
        let stream_path = self.object_presentation_path(key, index)?;
        let bytes = self
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        if bytes.as_ref() != patch.before_bytes() {
            return Err(OleError::InvalidFormat(
                "OLEDS presentation patch source does not match object stream".into(),
            ));
        }
        let source = PresentationSnapshot::parse_shared(bytes, patch.source().limits())?;
        let replacement = patch.apply(&source)?;
        self.put_stream_shared(&stream_path, replacement.bytes_shared())
    }

    /// Atomically edits one selected object's OLEDS native-data stream.
    ///
    /// Native bytes remain opaque and inert; the callback only receives the
    /// bounded stream transaction.
    pub fn update_native<F>(&mut self, key: &str, edit: F) -> Result<(), OleError>
    where
        F: FnOnce(&mut NativeTransaction) -> Result<(), OleError>,
    {
        self.update_native_with_limits(key, ole_streams::Limits::default(), edit)
    }

    /// Atomically edits one selected object's native-data stream under
    /// explicit stream limits.
    pub fn update_native_with_limits<F>(
        &mut self,
        key: &str,
        limits: ole_streams::Limits,
        edit: F,
    ) -> Result<(), OleError>
    where
        F: FnOnce(&mut NativeTransaction) -> Result<(), OleError>,
    {
        model::validate_ole_stream_limits(limits)?;
        let stream_path = self.object_native_path(key)?;
        let source = self
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)
            .and_then(|bytes| NativeSnapshot::parse_shared(bytes, limits))?;
        let mut transaction = source.edit();
        edit(&mut transaction)?;
        let commit = transaction.commit()?;
        self.put_stream_shared(&stream_path, commit.snapshot().bytes_shared())
    }

    /// Applies a source-checked OLEDS native-data patch to one selected object
    /// atomically.
    ///
    /// The exact source bytes are compared before reparsing, so a stale patch
    /// is rejected before an editor candidate is cloned or rendered.
    pub fn apply_native_patch(&mut self, key: &str, patch: &NativePatch) -> Result<(), OleError> {
        let stream_path = self.object_native_path(key)?;
        let bytes = self
            .package
            .stream_shared(&stream_path)
            .ok_or(OleError::StreamNotFound)?;
        if bytes.as_ref() != patch.before_bytes() {
            return Err(OleError::InvalidFormat(
                "OLEDS native patch source does not match object stream".into(),
            ));
        }
        let source = NativeSnapshot::parse_shared(bytes, patch.source().limits())?;
        let replacement = patch.apply(&source)?;
        self.put_stream_shared(&stream_path, replacement.bytes_shared())
    }

    fn object_presentation_path(&self, key: &str, index: usize) -> Result<Vec<String>, OleError> {
        let object = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?;
        object.validate_presentation_count()?;
        let mut path = object.path().to_vec();
        path.push(ole_streams::presentation_name(index)?);
        Ok(path)
    }

    fn object_native_path(&self, key: &str) -> Result<Vec<String>, OleError> {
        let object = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?;
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
    pub fn add_storage(&mut self, target: Target, compound_file: Vec<u8>) -> Result<(), OleError> {
        if compound_file.len() as u64 > self.limits.max_object_size {
            return Err(OleError::InvalidFormat(
                "new object exceeds size limit".into(),
            ));
        }
        let mut nested_ole = OleFile::open(Cursor::new(compound_file))?;
        codec::open(
            &nested_ole,
            self.limits
                .max_object_storages()
                .saturating_add(self.limits.max_object_streams()),
        )?;
        let nested = Package::capture_object(&mut nested_ole, self.limits)?;
        let mut candidate = self.clone();
        candidate
            .package
            .add_object(&target, &nested, self.limits)?;
        candidate.targets.push(target)?;
        *self = candidate.commit_candidate()?;
        Ok(())
    }

    /// Removes a selected storage after the host has removed its references.
    ///
    /// # Errors
    ///
    /// Returns an error when the target is missing or the rendered package
    /// fails validation.
    pub fn remove_storage(&mut self, key: &str) -> Result<Arc<[u8]>, OleError> {
        let object = self
            .objects
            .get(key)
            .ok_or_else(|| OleError::InvalidFormat(format!("object target {key:?} not found")))?;
        let removed = object.compound_shared();
        let mut candidate = self.clone();
        candidate
            .package
            .remove_object(object.path(), self.limits)?;
        candidate.targets = candidate.targets.without(key)?;
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
    pub const fn set_sector_layout_policy(&mut self, policy: SectorLayoutPolicy) {
        self.layout = policy;
    }

    /// The sector-layout policy a render will apply.
    #[must_use]
    pub const fn sector_layout_policy(&self) -> SectorLayoutPolicy {
        self.layout
    }

    /// Finishes the edit, returning the original bytes for a true no-op.
    ///
    /// # Errors
    ///
    /// Returns an error if the edited CFB cannot be rendered.
    pub fn finish(self) -> Result<Vec<u8>, OleError> {
        if self.changed {
            if self.layout == SectorLayoutPolicy::Reuse
                && let Some(rendered) = self.package.render_copy_through(
                    &self.base_package,
                    &self.original,
                    self.limits,
                )?
            {
                return Ok(rendered);
            }
            self.package
                .render_with_layout(Some(self.original.as_slice()), self.layout)
        } else {
            Ok(match Arc::try_unwrap(self.original) {
                Ok(bytes) => bytes,
                Err(bytes) => bytes.as_ref().clone(),
            })
        }
    }

    /// Commits this edit as an immutable snapshot plus a reversible patch.
    ///
    /// The source editor is consumed, so callers cannot accidentally keep
    /// mutating a value after using the commit result. The patch is checked
    /// against the exact original artifact and the snapshot has already
    /// passed the common CFB/resource validation performed by each edit.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited package cannot be rendered.
    pub fn commit(self) -> Result<Commit, OleError> {
        let before = self.original.as_ref().clone();
        let snapshot = self.snapshot();
        let after = self.finish()?;
        Ok(Commit::new(snapshot, Patch::new(before, after)))
    }

    fn commit_candidate(self) -> Result<Self, OleError> {
        self.commit_candidate_with_rendered()
            .map(|(candidate, _rendered)| candidate)
    }

    fn commit_candidate_with_rendered(mut self) -> Result<(Self, Vec<u8>), OleError> {
        self.package.check(self.limits)?;
        let rendered = if self.layout == SectorLayoutPolicy::Reuse {
            self.package
                .render_copy_through(&self.base_package, &self.original, self.limits)?
                .map_or_else(
                    || {
                        self.package
                            .render_with_layout(Some(self.original.as_slice()), self.layout)
                    },
                    Ok,
                )?
        } else {
            self.package
                .render_with_layout(Some(self.original.as_slice()), self.layout)?
        };
        let mut check = OleFile::open(Cursor::new(rendered.as_slice()))?;
        codec::open(&check, self.limits.max_package_directory_entries())?;
        let mut parsed = Package::capture(&mut check, self.limits)?;
        parsed.reuse_stream_allocations(&self.package)?;
        self.package = parsed;
        self.objects = discovery::from_package(&self.package, &self.targets, self.limits)?;
        self.changed = true;
        Ok((self, rendered))
    }
}
