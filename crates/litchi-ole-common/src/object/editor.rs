//! Transactional CFB stream and selected-storage editing.

use super::cfb_path::CfbPath;
use super::codec::{self, Package};
use super::discovery;
use super::link::{self, Link};
use super::model::{Limits, Objects};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::target::{Target, Targets};
use litchi_cfb::{OleError, OleFile};
use std::io::Cursor;
use std::sync::Arc;

/// Maximum number of stream selectors accepted by one removal publication.
pub const MAX_STREAM_REMOVALS: usize = 1_024;

/// Transactional editor for target-selected OLE storages.
#[derive(Debug, Clone)]
pub struct Editor {
    targets: Targets,
    limits: Limits,
    original: Arc<Vec<u8>>,
    package: Package,
    objects: Objects,
    changed: bool,
}

impl Editor {
    /// Opens a CFB package with an explicit host-resolved target catalog.
    ///
    /// # Errors
    ///
    /// Returns an error when the CFB is malformed or protected, a target is
    /// missing/invalid, or a configured resource limit is exceeded.
    pub fn open(bytes: Vec<u8>, targets: Targets, limits: Limits) -> Result<Self, OleError> {
        limits.validate()?;
        if targets.len() > limits.max_objects {
            return Err(OleError::InvalidFormat(format!(
                "object target count {} exceeds limit {}",
                targets.len(),
                limits.max_objects
            )));
        }
        let original = Arc::new(bytes);
        let mut ole = OleFile::open(Cursor::new(original.as_slice()))?;
        codec::open(&ole)?;
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
            .map(|target| target.resolve(&ole))
            .collect::<Result<Vec<_>, _>>()?;
        let resolved_targets = Targets::new(resolved_target_entries)?;
        let package = Package::capture(&mut ole, limits)?;
        package.check(limits)?;
        let objects = discovery::from_package(&package, &resolved_targets, limits)?;
        Ok(Self {
            targets: resolved_targets,
            limits,
            original,
            package,
            objects,
            changed: false,
        })
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
            self.package.clone(),
            self.objects.clone(),
            self.changed,
        )
    }

    pub(crate) fn from_snapshot(snapshot: &Snapshot) -> Self {
        Self {
            targets: snapshot.targets().clone(),
            limits: snapshot.limits(),
            original: snapshot.original(),
            package: snapshot.package(),
            objects: snapshot.objects_clone(),
            changed: snapshot.changed(),
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
            return Ok(());
        }
        let mut replacement_ole = OleFile::open(Cursor::new(compound_file))?;
        codec::open(&replacement_ole)?;
        let replacement = Package::capture(&mut replacement_ole, self.limits)?;
        let mut candidate = self.clone();
        candidate
            .package
            .replace_object(object.path(), &replacement, self.limits)?;
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
        codec::open(&nested_ole)?;
        let nested = Package::capture(&mut nested_ole, self.limits)?;
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

    /// Finishes the edit, returning the original bytes for a true no-op.
    ///
    /// # Errors
    ///
    /// Returns an error if the edited CFB cannot be rendered.
    pub fn finish(self) -> Result<Vec<u8>, OleError> {
        if self.changed {
            self.package.render()
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
        let rendered = self.package.render()?;
        let mut check = OleFile::open(Cursor::new(rendered.as_slice()))?;
        codec::open(&check)?;
        let mut parsed = Package::capture(&mut check, self.limits)?;
        parsed.reuse_stream_allocations(&self.package)?;
        self.package = parsed;
        self.objects = discovery::from_package(&self.package, &self.targets, self.limits)?;
        self.changed = true;
        Ok((self, rendered))
    }
}
