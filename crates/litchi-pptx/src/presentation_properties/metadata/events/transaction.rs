//! Detached source-checked edits for one slide's inert show-event list.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec::{self, LoadLimits, Located};
use super::model::{Draft, Event, Kind};
use super::validation::{validate_draft, validate_events};
use crate::time::Offset;
use crate::{Error, Result};

/// Stable fingerprint of the exact owning slide XML bytes.
pub type Revision = u64;

/// An immutable semantic view bound to one exact slide source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) slide_part_name: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) events: Vec<Event>,
    pub(crate) layout: Located,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load the slide's persisted show-event metadata.
    pub fn load(
        package: &OpcPackage,
        slide_part_name: &litchi_opc::PackURI,
    ) -> Result<Option<Self>> {
        super::package::load_snapshot(package, slide_part_name)
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(
        package: &OpcPackage,
        slide_part_name: &litchi_opc::PackURI,
    ) -> Result<Option<Self>> {
        Self::load(package, slide_part_name)
    }

    pub(crate) fn from_located(
        slide_part_name: String,
        source_xml: Arc<Vec<u8>>,
        layout: Located,
    ) -> Result<Self> {
        let drafts: Vec<_> = layout.events.iter().map(Draft::from_event).collect();
        validate_events(&drafts)?;
        Ok(Self {
            slide_part_name,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            events: layout.events.clone(),
            layout,
        })
    }

    /// Return the owning slide part name.
    #[inline]
    pub fn slide_part_name(&self) -> &str {
        &self.slide_part_name
    }

    /// Borrow the ordered typed event records.
    #[inline]
    pub fn events(&self) -> &[Event] {
        &self.events
    }

    /// Return the source fingerprint used for stale-source checks.
    #[inline]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact owning slide XML captured by this snapshot.
    #[inline]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit over the typed event sequence.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.events.iter().map(Draft::from_event).collect(),
        }
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.slide_part_name == other.slide_part_name
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.revision == other.revision
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }
}

/// A bounded edit staged against one existing event list.
#[derive(Clone, Debug)]
pub struct Transaction {
    original: Snapshot,
    working: Vec<Draft>,
}

impl Transaction {
    /// Borrow the projected typed draft sequence.
    #[inline]
    pub fn events(&self) -> &[Draft] {
        &self.working
    }

    /// Alias for [`Self::events`].
    #[inline]
    pub fn snapshot(&self) -> &[Draft] {
        self.events()
    }

    /// Return whether the staged semantic sequence differs from the source.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.working
            != self
                .original
                .events
                .iter()
                .map(Draft::from_event)
                .collect::<Vec<_>>()
    }

    /// Replace all typed records while retaining the owning extension bytes.
    pub fn replace(&mut self, events: Vec<Draft>) -> Result<bool> {
        validate_events(&events)?;
        if self.working == events {
            return Ok(false);
        }
        self.working = events;
        Ok(true)
    }

    /// Apply a checked mutation to one event without partially staging failure.
    pub fn edit_event(
        &mut self,
        index: usize,
        edit: impl FnOnce(&mut Draft) -> Result<()>,
    ) -> Result<()> {
        let mut candidate = self
            .working
            .get(index)
            .cloned()
            .ok_or_else(|| index_error(index, self.working.len()))?;
        edit(&mut candidate)?;
        validate_draft(&candidate)?;
        let mut working = self.working.clone();
        working[index] = candidate;
        validate_events(&working)?;
        self.working = working;
        Ok(())
    }

    /// Set one event's complete typed kind.
    pub fn set_kind(&mut self, index: usize, kind: Kind) -> Result<()> {
        self.edit_event(index, |event| {
            event.set_kind(kind);
            Ok(())
        })
    }

    /// Set one event's timeline offset.
    pub fn set_time(&mut self, index: usize, time: Offset) -> Result<()> {
        self.edit_event(index, |event| {
            event.set_time(time);
            Ok(())
        })
    }

    /// Set one event's DrawingML object identifier.
    pub fn set_object_id(&mut self, index: usize, object_id: u32) -> Result<()> {
        self.edit_event(index, |event| {
            event.set_object_id(object_id);
            Ok(())
        })
    }

    /// Append one event to the ordered list.
    pub fn push(&mut self, event: Draft) -> Result<()> {
        let mut candidate = self.working.clone();
        candidate.push(event);
        validate_events(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Insert one event before an existing source-order index.
    pub fn insert(&mut self, index: usize, event: Draft) -> Result<()> {
        if index > self.working.len() {
            return Err(index_error(index, self.working.len()));
        }
        let mut candidate = self.working.clone();
        candidate.insert(index, event);
        validate_events(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Remove one event, retaining the non-empty-list invariant.
    pub fn remove(&mut self, index: usize) -> Result<Draft> {
        if self.working.len() <= 1 {
            return Err(invalid(
                "cannot remove the last slide-show event; remove the package extension instead",
            ));
        }
        if index >= self.working.len() {
            return Err(index_error(index, self.working.len()));
        }
        let mut candidate = self.working.clone();
        let removed = candidate.remove(index);
        validate_events(&candidate)?;
        self.working = candidate;
        Ok(removed)
    }

    /// Move one event to another source-order position.
    pub fn move_event(&mut self, from: usize, to: usize) -> Result<bool> {
        if from >= self.working.len() || to >= self.working.len() {
            return Err(index_error(from.max(to), self.working.len()));
        }
        if from == to {
            return Ok(false);
        }
        let mut candidate = self.working.clone();
        let event = candidate.remove(from);
        candidate.insert(to, event);
        validate_events(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Consume the detached edit into a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        validate_events(&self.working)?;
        let updated = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.layout,
            &self.working,
        )?;
        let updated = Arc::new(updated);
        let slide_index = self.original.events.first().map_or(0, Event::slide_index);
        let located = codec::locate(slide_index, updated.as_slice(), &mut LoadLimits::default())?
            .ok_or_else(|| invalid("edited slide-show event list disappeared"))?;
        let parsed: Vec<_> = located.events.iter().map(Draft::from_event).collect();
        if parsed != self.working {
            return Err(invalid(
                "slide-show event serialization changed the typed sequence",
            ));
        }
        let snapshot = Snapshot::from_located(
            self.original.slide_part_name.clone(),
            Arc::clone(&updated),
            located,
        )?;
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful event edit and its reversible package patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether publication changes the exact slide bytes.
    #[inline]
    pub const fn changed(&self) -> bool {
        self.changed
    }

    /// Alias for [`Self::changed`].
    #[inline]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its patch.
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible source-checked replacement of one slide's event list.
#[derive(Clone, Debug)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before publication.
    #[inline]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by publication.
    #[inline]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Alias for [`Self::is_empty`].
    #[inline]
    pub fn is_changed(&self) -> bool {
        !self.is_empty()
    }

    /// Return the exact inverse patch.
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return the source fingerprint required for publication.
    #[inline]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Apply this patch atomically after checking the complete slide source.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }
}

fn fingerprint(bytes: &[u8]) -> Revision {
    let mut hash = 0xcbf29ce484222325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100000001b3);
    }
    hash
}

fn index_error(index: usize, len: usize) -> Error {
    invalid(format!(
        "slide-show event index {index} is outside a list of length {len}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
