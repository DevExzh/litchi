//! Source-checked semantic edits for the external-media owner boundary.

use std::sync::Arc;

use crate::consts::RecordType;
use crate::package::{Error, Result};

use super::model::{Collection, Limits, LinkedAudio, Movie, Object, Playback};
use super::{package, validation};

/// Compact identity of one exact serialized PPT root record.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct Revision(u64);

impl Revision {
    fn from_bytes(bytes: &[u8]) -> Self {
        let mut value = 0xcbf2_9ce4_8422_2325u64;
        value ^= bytes.len() as u64;
        value = value.wrapping_mul(0x0000_0100_0000_01b3);
        for byte in bytes {
            value ^= u64::from(*byte);
            value = value.wrapping_mul(0x0000_0100_0000_01b3);
        }
        Self(value)
    }

    /// Compact source fingerprint suitable for owner conflict checks.
    #[must_use]
    pub const fn value(self) -> u64 {
        self.0
    }
}

/// One typed semantic operation staged by an external-media transaction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Change {
    /// Change or clear the inert path of a movie or linked-audio object.
    SetPath {
        id: u32,
        before: Option<String>,
        after: Option<String>,
    },
    /// Change the flags carried by an `ExMediaAtom`.
    SetPlayback {
        id: u32,
        before: Playback,
        after: Playback,
    },
    /// Add a new media definition to the owning list.
    Insert { object: Object },
    /// Replace a media definition while retaining its owner ID.
    Replace {
        id: u32,
        before: Object,
        after: Object,
    },
    /// Remove an unowned media definition.
    Remove { object: Object },
}

impl Change {
    fn inverse(&self) -> Self {
        match self {
            Self::SetPath { id, before, after } => Self::SetPath {
                id: *id,
                before: after.clone(),
                after: before.clone(),
            },
            Self::SetPlayback { id, before, after } => Self::SetPlayback {
                id: *id,
                before: *after,
                after: *before,
            },
            Self::Insert { object } => Self::Remove {
                object: object.clone(),
            },
            Self::Replace { id, before, after } => Self::Replace {
                id: *id,
                before: after.clone(),
                after: before.clone(),
            },
            Self::Remove { object } => Self::Insert {
                object: object.clone(),
            },
        }
    }
}

/// Immutable, source-preserving view of one complete PPT root record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    bytes: Arc<[u8]>,
    collection: Option<Collection>,
    location: Option<package::Location>,
    root_type: RecordType,
    owner_ids: Vec<u32>,
    hyperlink_ids: Vec<u32>,
    revision: Revision,
    limits: Limits,
}

impl Snapshot {
    /// Parse one complete root record with default limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: impl AsRef<[u8]>) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one complete root record under explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: impl AsRef<[u8]>, limits: Limits) -> Result<Self> {
        Self::from_bytes_with_limits(bytes.as_ref().to_vec(), limits)
    }

    /// Capture an owned root record without another caller-side copy.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Capture an owned root record under explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        let validated = validation::parse_source(&bytes, limits)?;
        let shared_bytes: Arc<[u8]> = Arc::from(bytes.into_boxed_slice());
        let revision = Revision::from_bytes(&shared_bytes);
        Ok(Self {
            bytes: shared_bytes,
            collection: validated.collection,
            location: validated.location,
            root_type: validated.root_type,
            owner_ids: validated.owner_ids,
            hyperlink_ids: validated.hyperlink_ids,
            revision,
            limits,
        })
    }

    /// Exact source or committed root-record bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Compact identity of the exact serialized source.
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Resource limits retained for subsequent edits.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Typed external-media list, if the root owns one.
    #[must_use]
    pub fn collection(&self) -> Option<&Collection> {
        self.collection.as_ref()
    }

    /// IDs referenced by parsed `ExObjRefAtom` owners in this root tree.
    #[must_use]
    pub fn owner_ids(&self) -> &[u32] {
        &self.owner_ids
    }

    /// Number of parsed owners for one media ID.
    #[must_use]
    pub fn owner_count(&self, id: u32) -> usize {
        self.owner_ids.iter().filter(|owner| **owner == id).count()
    }

    /// Begin an isolated semantic edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            candidate: self.collection.clone(),
            changes: Vec::new(),
        }
    }
}

/// Isolated, failure-atomic editor over one source snapshot.
#[derive(Clone, Debug)]
pub struct Transaction {
    source: Snapshot,
    candidate: Option<Collection>,
    changes: Vec<Change>,
}

impl Transaction {
    /// Immutable source snapshot used for stale-source checks.
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.source
    }

    /// Current typed media collection, if present.
    #[must_use]
    pub fn collection(&self) -> Option<&Collection> {
        self.candidate.as_ref()
    }

    /// Current typed media objects, or an empty slice when no list exists.
    #[must_use]
    pub fn objects(&self) -> &[Object] {
        self.candidate
            .as_ref()
            .map_or(&[], |collection| collection.objects.as_slice())
    }

    /// Whether staged semantics differ from the source collection.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.candidate != self.source.collection
    }

    /// Staged semantic operations in source order.
    #[must_use]
    pub fn changes(&self) -> &[Change] {
        &self.changes
    }

    /// Change or clear an inert movie/audio path without resolving it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_path(&mut self, id: u32, path: Option<String>) -> Result<()> {
        let mut candidate = self.require_collection()?.clone();
        let object = candidate
            .objects
            .iter_mut()
            .find(|object| object.id() == id)
            .ok_or_else(|| Error::InvalidFormat(format!("external media ID {id} was not found")))?;
        let before = path_of(object)?;
        if before == path {
            return Ok(());
        }
        set_path_of(object, path.clone())?;
        self.validate_candidate(&candidate)?;
        self.candidate = Some(candidate);
        self.changes.push(Change::SetPath {
            id,
            before,
            after: path,
        });
        Ok(())
    }

    /// Change the inert playback flags of one media object.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_playback(&mut self, id: u32, playback: Playback) -> Result<()> {
        let mut candidate = self.require_collection()?.clone();
        let object = candidate
            .objects
            .iter_mut()
            .find(|object| object.id() == id)
            .ok_or_else(|| Error::InvalidFormat(format!("external media ID {id} was not found")))?;
        let before = object.playback();
        if before == playback {
            return Ok(());
        }
        set_playback_of(object, playback);
        self.validate_candidate(&candidate)?;
        self.candidate = Some(candidate);
        self.changes.push(Change::SetPlayback {
            id,
            before,
            after: playback,
        });
        Ok(())
    }

    /// Insert a typed media definition, allocating no external content.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert(&mut self, object: Object) -> Result<()> {
        let id = object.id();
        if id == 0 {
            return Err(Error::InvalidFormat(
                "external media IDs must be positive".into(),
            ));
        }
        if validation::id_is_reserved(id, &self.source.hyperlink_ids, &self.source.owner_ids) {
            return Err(Error::InvalidFormat(format!(
                "external media ID {id} is occupied by a hyperlink or owner"
            )));
        }
        let mut candidate = self.candidate.clone().unwrap_or(Collection {
            id_seed: id,
            objects: Vec::new(),
            unknown_records: Vec::new(),
        });
        if candidate.get(id).is_some() {
            return Err(Error::InvalidFormat(format!(
                "external media ID {id} is already present"
            )));
        }
        candidate.id_seed = candidate.id_seed.max(id);
        candidate.objects.push(object.clone());
        self.validate_candidate(&candidate)?;
        self.candidate = Some(candidate);
        self.changes.push(Change::Insert { object });
        Ok(())
    }

    /// Replace a definition while retaining the same owner ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, id: u32, replacement: Object) -> Result<Object> {
        if replacement.id() != id {
            return Err(Error::InvalidFormat(
                "external media replacement must retain its owner ID".into(),
            ));
        }
        let mut candidate = self.require_collection()?.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| Error::InvalidFormat(format!("external media ID {id} was not found")))?;
        let before = candidate.objects[index].clone();
        if before == replacement {
            return Ok(before);
        }
        candidate.objects[index] = replacement.clone();
        self.validate_candidate(&candidate)?;
        self.candidate = Some(candidate);
        self.changes.push(Change::Replace {
            id,
            before: before.clone(),
            after: replacement,
        });
        Ok(before)
    }

    /// Remove a media definition only when no parsed owner still points at it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, id: u32) -> Result<Object> {
        let current = self.require_collection()?;
        validation::can_remove(current, &self.source.owner_ids, id)?;
        let mut candidate = current.clone();
        let index = candidate
            .objects
            .iter()
            .position(|object| object.id() == id)
            .ok_or_else(|| Error::InvalidFormat(format!("external media ID {id} was not found")))?;
        let object = candidate.objects.remove(index);
        for record in &mut candidate.unknown_records {
            if record.object_index > index {
                record.object_index -= 1;
            }
        }
        self.validate_candidate(&candidate)?;
        self.candidate = Some(candidate);
        self.changes.push(Change::Remove {
            object: object.clone(),
        });
        Ok(object)
    }

    /// Capture the current candidate without publishing it.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn snapshot(&self) -> Result<Snapshot> {
        if !self.is_changed() {
            return Ok(self.source.clone());
        }
        Snapshot::from_bytes_with_limits(self.emit_candidate()?, self.source.limits)
    }

    /// Publish the candidate atomically with a reversible source-checked patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch {
                base: self.source.revision,
                target: self.source.revision,
                before: self.source.bytes.clone(),
                after: self.source.bytes.clone(),
                changes: Vec::new(),
                limits: self.source.limits,
            };
            return Ok(Commit {
                snapshot: self.source,
                patch,
            });
        }
        let bytes = self.emit_candidate()?;
        let source = self.source;
        let changes = self.changes;
        let snapshot = Snapshot::from_bytes_with_limits(bytes, source.limits)?;
        let patch = Patch {
            base: source.revision,
            target: snapshot.revision,
            before: source.bytes,
            after: snapshot.bytes.clone(),
            changes,
            limits: snapshot.limits,
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

    fn require_collection(&self) -> Result<&Collection> {
        self.candidate
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("the root has no external-media collection".into()))
    }

    fn validate_candidate(&self, candidate: &Collection) -> Result<()> {
        validation::validate_collection(candidate, &self.source.hyperlink_ids)
    }

    fn emit_candidate(&self) -> Result<Vec<u8>> {
        let candidate = self.candidate.as_ref().ok_or_else(|| {
            Error::InvalidFormat("the root has no external-media collection".into())
        })?;
        self.validate_candidate(candidate)?;
        let replacement = candidate.to_record_bytes()?;
        match &self.source.location {
            Some(location) => package::replace(&self.source.bytes, location, &replacement),
            None => {
                package::append_to_document(&self.source.bytes, self.source.root_type, &replacement)
            },
        }
    }
}

/// A successful immutable target and its source-checked patch.
#[derive(Clone, Debug, PartialEq, Eq)]
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

/// A source-checked reversible patch for one complete PPT root record.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Patch {
    base: Revision,
    target: Revision,
    before: Arc<[u8]>,
    after: Arc<[u8]>,
    changes: Vec<Change>,
    limits: Limits,
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

    /// Exact target bytes produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Whether the patch is an exact byte-for-byte no-op.
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
                "external-media patch source does not match its base snapshot".into(),
            ));
        }
        if self.is_empty() {
            return Ok(current.clone());
        }
        Snapshot::from_bytes_with_limits(self.after.to_vec(), self.limits)
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
            limits: self.limits,
        }
    }
}

fn path_of(object: &Object) -> Result<Option<String>> {
    match object {
        Object::Movie(Movie { video, .. }) => Ok(video.path.clone()),
        Object::LinkedAudio(LinkedAudio { path, .. }) => Ok(path.clone()),
        Object::CdAudio(_) | Object::EmbeddedWav(_) => Err(Error::InvalidFormat(
            "CD and embedded WAV media do not carry external paths".into(),
        )),
    }
}

fn set_path_of(object: &mut Object, path: Option<String>) -> Result<()> {
    match object {
        Object::Movie(value) => value.video.path = path,
        Object::LinkedAudio(value) => value.path = path,
        Object::CdAudio(_) | Object::EmbeddedWav(_) => {
            return Err(Error::InvalidFormat(
                "CD and embedded WAV media do not carry external paths".into(),
            ));
        },
    }
    Ok(())
}

fn set_playback_of(object: &mut Object, playback: Playback) {
    let media = match object {
        Object::Movie(value) => &mut value.video.media,
        Object::LinkedAudio(value) => &mut value.media,
        Object::CdAudio(value) => &mut value.media,
        Object::EmbeddedWav(value) => &mut value.media,
    };
    media.loop_playback = playback.loop_playback;
    media.rewind_after_playing = playback.rewind_after_playing;
    media.narration = playback.narration;
}
