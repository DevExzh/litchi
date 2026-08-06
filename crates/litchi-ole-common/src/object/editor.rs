//! Transactional CFB stream and selected-storage editing.

use super::codec::{self, Package};
use super::discovery;
use super::model::{Limits, Objects};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::target::{Target, Targets};
use litchi_cfb::{OleError, OleFile};
use std::io::Cursor;
use std::sync::Arc;

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
        let targets = targets
            .as_slice()
            .iter()
            .map(|target| target.resolve(&ole))
            .collect::<Result<Vec<_>, _>>()?;
        let targets = Targets::new(targets)?;
        let package = Package::capture(&mut ole, limits)?;
        package.check(limits)?;
        let objects = discovery::from_package(&package, &targets, limits)?;
        Ok(Self {
            targets,
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
    pub fn commit(self) -> Result<Commit, OleError> {
        let before = self.original.as_ref().clone();
        let snapshot = self.snapshot();
        let after = self.finish()?;
        Ok(Commit::new(snapshot, Patch::new(before, after)))
    }

    fn commit_candidate(mut self) -> Result<Self, OleError> {
        self.package.check(self.limits)?;
        let rendered = self.package.render()?;
        let mut check = OleFile::open(Cursor::new(rendered))?;
        codec::open(&check)?;
        let parsed = Package::capture(&mut check, self.limits)?;
        self.objects = discovery::from_package(&parsed, &self.targets, self.limits)?;
        self.changed = true;
        Ok(self)
    }
}
