//! Detached source-checked edits for presentation extended guides.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec;
use super::model::{Guide, Guides, List, ListKind};
use crate::{Error, Result};

/// Stable fingerprint of the exact owning presentation XML bytes.
pub type Revision = u64;

/// An immutable typed view bound to one exact presentation source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) presentation_part_name: String,
    pub(crate) presentation_content_type: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) guides: Guides,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load the presentation's extended-guide owner from an OPC package.
    pub fn load(package: &OpcPackage) -> Result<Self> {
        super::package::load_snapshot(package)
    }

    /// Alias for [`Self::load`] emphasizing the source-bound result.
    pub fn read(package: &OpcPackage) -> Result<Self> {
        Self::load(package)
    }

    /// Parse an exact presentation XML source into a detached snapshot.
    pub fn from_xml(source_xml: &[u8]) -> Result<Self> {
        Self::from_wire(String::new(), String::new(), Arc::new(source_xml.to_vec()))
    }

    pub(crate) fn from_wire(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
    ) -> Result<Self> {
        if source_xml.len() > codec::MAX_BYTES {
            return Err(invalid("presentation-guide source exceeds 8 MiB"));
        }
        let guides = Guides::from_xml(source_xml.as_slice())?;
        Ok(Self {
            presentation_part_name,
            presentation_content_type,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            guides,
        })
    }

    /// Borrow the complete typed extended-guide value.
    #[inline]
    pub fn guides(&self) -> &Guides {
        &self.guides
    }

    /// Alias emphasizing the semantic guide catalog.
    #[inline]
    pub fn value(&self) -> &Guides {
        self.guides()
    }

    /// Return the owning presentation part name.
    #[inline]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the source presentation content type.
    #[inline]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Return the fingerprint used for stale-source checks.
    #[inline]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact presentation XML captured by this snapshot.
    #[inline]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit over the typed guide value.
    #[inline]
    pub fn edit(&self) -> Transaction {
        Transaction {
            original: self.clone(),
            working: self.guides.clone(),
        }
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.presentation_content_type == other.presentation_content_type
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.revision == other.revision
    }
}

/// A failure-atomic edit staged against one guide catalog.
#[derive(Clone, Debug)]
pub struct Transaction {
    original: Snapshot,
    working: Guides,
}

impl Transaction {
    /// Borrow the immutable source snapshot used by this edit.
    #[inline]
    pub const fn source(&self) -> &Snapshot {
        &self.original
    }

    /// Borrow the currently staged guide value.
    #[inline]
    pub fn guides(&self) -> &Guides {
        &self.working
    }

    /// Alias for [`Self::guides`].
    #[inline]
    pub fn snapshot(&self) -> &Guides {
        self.guides()
    }

    /// Borrow one currently staged owner list.
    #[inline]
    pub fn list(&self, kind: ListKind) -> Option<&List> {
        self.working.list(kind)
    }

    /// Return whether the staged typed value differs from the source.
    #[inline]
    pub fn is_changed(&self) -> bool {
        self.original.guides != self.working
    }

    /// Replace both owner lists after validating IDs, bounds, and opaque XML.
    pub fn replace(&mut self, value: Guides) -> Result<bool> {
        codec::validate_value(&value)?;
        if self.working == value {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to the complete guide value atomically.
    pub fn edit(&mut self, operation: impl FnOnce(&mut Guides) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        operation(&mut candidate)?;
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Replace one slide or notes guide list, retaining all other metadata.
    pub fn set_list(&mut self, kind: ListKind, value: Option<List>) -> Result<bool> {
        let mut candidate = self.working.clone();
        set_list(&mut candidate, kind, value);
        if candidate == self.working {
            return Ok(false);
        }
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Apply a checked mutation to one existing owner list atomically.
    pub fn edit_list(
        &mut self,
        kind: ListKind,
        operation: impl FnOnce(&mut List) -> Result<()>,
    ) -> Result<()> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        operation(list)?;
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Append one guide, creating the selected owner list when needed.
    pub fn push(&mut self, kind: ListKind, guide: Guide) -> Result<()> {
        self.insert(
            kind,
            self.list(kind).map_or(0, |list| list.guides.len()),
            guide,
        )
    }

    /// Insert one guide before a checked source-order index.
    pub fn insert(&mut self, kind: ListKind, index: usize, guide: Guide) -> Result<()> {
        let mut candidate = self.working.clone();
        let list = list_mut_or_create(&mut candidate, kind);
        if index > list.guides.len() {
            return Err(index_error(index, list.guides.len()));
        }
        list.guides.insert(index, guide);
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Replace one guide while preserving its source-order position.
    pub fn replace_guide(&mut self, kind: ListKind, index: usize, guide: Guide) -> Result<bool> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        let length = list.guides.len();
        let current = list
            .guides
            .get_mut(index)
            .ok_or_else(|| index_error(index, length))?;
        if *current == guide {
            return Ok(false);
        }
        *current = guide;
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Apply a checked mutation to one guide without partial staging.
    pub fn edit_guide(
        &mut self,
        kind: ListKind,
        index: usize,
        operation: impl FnOnce(&mut Guide) -> Result<()>,
    ) -> Result<()> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        let length = list.guides.len();
        let guide = list
            .guides
            .get_mut(index)
            .ok_or_else(|| index_error(index, length))?;
        operation(guide)?;
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Change one guide's native ID after checking list uniqueness.
    pub fn set_id(&mut self, kind: ListKind, index: usize, id: u32) -> Result<()> {
        self.edit_guide(kind, index, |guide| {
            guide.id = id;
            Ok(())
        })
    }

    /// Remove one guide by source-order index.
    pub fn remove(&mut self, kind: ListKind, index: usize) -> Result<Guide> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        if index >= list.guides.len() {
            return Err(index_error(index, list.guides.len()));
        }
        let removed = list.guides.remove(index);
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(removed)
    }

    /// Remove one guide by its validated native ID.
    pub fn remove_id(&mut self, kind: ListKind, id: u32) -> Result<Option<Guide>> {
        let Some(list) = self.list(kind) else {
            return Ok(None);
        };
        let Some(index) = list.guides.iter().position(|guide| guide.id == id) else {
            return Ok(None);
        };
        self.remove(kind, index).map(Some)
    }

    /// Move one guide to another checked position in the same owner list.
    pub fn move_guide(&mut self, kind: ListKind, from: usize, to: usize) -> Result<bool> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        if from >= list.guides.len() || to >= list.guides.len() {
            return Err(index_error(from.max(to), list.guides.len()));
        }
        if from == to {
            return Ok(false);
        }
        let guide = list.guides.remove(from);
        list.guides.insert(to, guide);
        codec::validate_value(&candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Clear one owner list while retaining the other list and its source data.
    pub fn clear(&mut self, kind: ListKind) -> Result<bool> {
        self.set_list(kind, None)
    }

    /// Validate and consume this edit into a source-checked commit.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.original.clone(), self.original.clone());
            return Ok(Commit {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        codec::validate_value(&self.working)?;
        let updated = Arc::new(codec::rewrite_source(
            self.original.source_xml(),
            &self.working,
        )?);
        let guides = Guides::from_xml(updated.as_slice())?;
        if guides != self.working {
            return Err(invalid(
                "extended-guide serialization changed the typed value",
            ));
        }
        let snapshot = Snapshot::from_wire(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            updated,
        )?;
        let patch = Patch::new(self.original, snapshot.clone());
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful guide edit and its reversible package patch.
#[derive(Clone, Debug)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether publication changes the exact presentation bytes.
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

/// A reversible source-checked replacement of one presentation XML source.
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

    /// Alias for [`Self::is_empty`] with mutation-oriented naming.
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

    /// Apply this patch atomically to an OPC package.
    pub fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }

    /// Apply this patch atomically to an owned presentation XML buffer.
    pub fn apply_xml(&self, target: &mut Vec<u8>) -> Result<Snapshot> {
        if target.as_slice() != self.before.source_xml() {
            return Err(invalid("extended-guide source is stale"));
        }
        if self.is_empty() {
            return Ok(self.before.clone());
        }
        let parsed = Snapshot::from_wire(
            self.after.presentation_part_name.clone(),
            self.after.presentation_content_type.clone(),
            Arc::clone(self.after.source_arc()),
        )?;
        if !parsed.same_source(&self.after) {
            return Err(invalid(
                "extended-guide patch output differs from its committed source",
            ));
        }
        target.clone_from(self.after.source_xml.as_ref());
        Ok(parsed)
    }
}

fn set_list(value: &mut Guides, kind: ListKind, list: Option<List>) {
    match kind {
        ListKind::Slide => value.slide = list,
        ListKind::Notes => value.notes = list,
    }
}

fn list_mut(value: &mut Guides, kind: ListKind) -> Option<&mut List> {
    match kind {
        ListKind::Slide => value.slide.as_mut(),
        ListKind::Notes => value.notes.as_mut(),
    }
}

fn list_mut_or_create(value: &mut Guides, kind: ListKind) -> &mut List {
    match kind {
        ListKind::Slide => value.slide.get_or_insert_with(List::default),
        ListKind::Notes => value.notes.get_or_insert_with(List::default),
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
        "extended-guide index {index} is outside a list of length {len}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
