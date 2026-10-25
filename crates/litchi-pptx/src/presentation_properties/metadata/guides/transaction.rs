//! Detached source-checked edits for presentation extended guides.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::Limits;
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
    pub(crate) limits: Limits,
    pub(crate) revision: Revision,
}

impl Snapshot {
    /// Load the presentation's extended-guide owner from an OPC package.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load(package: &OpcPackage, limits: Limits) -> Result<Self> {
        super::package::load_snapshot(package, limits)
    }

    /// Parse an exact presentation XML source into a detached snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn from_xml(source_xml: &[u8], limits: Limits) -> Result<Self> {
        // Admit the borrowed source before retaining it.  In particular, do
        // not let an oversized caller buffer turn into an owned copy before
        // the selected input policy has had a chance to reject it.
        let guides = Guides::from_xml(source_xml, limits)?;
        let source_xml = Arc::new(source_xml.to_vec());
        Ok(Self::from_parsed(
            String::new(),
            String::new(),
            source_xml,
            guides,
            limits,
        ))
    }

    pub(crate) fn from_wire(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
        limits: Limits,
    ) -> Result<Self> {
        if source_xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                observed: source_xml.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        let guides = Guides::from_xml(source_xml.as_slice(), limits)?;
        Ok(Self::from_parsed(
            presentation_part_name,
            presentation_content_type,
            source_xml,
            guides,
            limits,
        ))
    }

    fn from_parsed(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
        guides: Guides,
        limits: Limits,
    ) -> Self {
        Self {
            presentation_part_name,
            presentation_content_type,
            revision: fingerprint(source_xml.as_slice()),
            source_xml,
            guides,
            limits,
        }
    }

    /// Borrow the complete typed extended-guide value.
    #[inline]
    #[must_use]
    pub fn guides(&self) -> &Guides {
        &self.guides
    }

    /// Return the owning presentation part name.
    #[inline]
    #[must_use]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the source presentation content type.
    #[inline]
    #[must_use]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Return the fingerprint used for stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Borrow the exact presentation XML captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Start an atomic detached edit over the typed guide value.
    #[inline]
    #[must_use]
    pub(crate) fn edit(&self) -> Draft {
        Draft {
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
pub struct Draft {
    original: Snapshot,
    working: Guides,
}

impl Draft {
    /// Borrow the immutable source snapshot used by this edit.
    #[inline]
    #[must_use]
    pub const fn source(&self) -> &Snapshot {
        &self.original
    }

    /// Borrow the currently staged guide value.
    #[inline]
    #[must_use]
    pub fn guides(&self) -> &Guides {
        &self.working
    }

    /// Borrow one currently staged owner list.
    #[inline]
    #[must_use]
    pub fn list(&self, kind: ListKind) -> Option<&List> {
        self.working.list(kind)
    }

    fn has_changes(&self) -> bool {
        self.original.guides != self.working
    }

    /// The outcome of this edit or commit.
    #[must_use]
    pub fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.has_changes())
    }

    /// Replace both owner lists after validating IDs, bounds, and opaque XML.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, value: Guides) -> Result<bool> {
        codec::validate_value(&value, self.original.limits)?;
        if self.working == value {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to the complete guide value atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit(&mut self, operation: impl FnOnce(&mut Guides) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        operation(&mut candidate)?;
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(())
    }

    /// Replace one slide or notes guide list, retaining all other metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_list(&mut self, kind: ListKind, value: Option<List>) -> Result<bool> {
        let mut candidate = self.working.clone();
        set_list(&mut candidate, kind, value);
        if candidate == self.working {
            return Ok(false);
        }
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(true)
    }

    /// Apply a checked mutation to one existing owner list atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit_list(
        &mut self,
        kind: ListKind,
        operation: impl FnOnce(&mut List) -> Result<()>,
    ) -> Result<()> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        operation(list)?;
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(())
    }

    /// Append one guide, creating the selected owner list when needed.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn push(&mut self, kind: ListKind, guide: Guide) -> Result<()> {
        self.insert(
            kind,
            self.list(kind).map_or(0, |list| list.guides.len()),
            guide,
        )
    }

    /// Insert one guide before a checked source-order index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert(&mut self, kind: ListKind, index: usize, guide: Guide) -> Result<()> {
        let mut candidate = self.working.clone();
        let list = list_mut_or_create(&mut candidate, kind);
        if index > list.guides.len() {
            return Err(index_error(index, list.guides.len()));
        }
        list.guides.insert(index, guide);
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(())
    }

    /// Replace one guide while preserving its source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(true)
    }

    /// Apply a checked mutation to one guide without partial staging.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(())
    }

    /// Change one guide's native ID after checking list uniqueness.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_id(&mut self, kind: ListKind, index: usize, id: u32) -> Result<()> {
        self.edit_guide(kind, index, |guide| {
            guide.id = id;
            Ok(())
        })
    }

    /// Remove one guide by source-order index.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, kind: ListKind, index: usize) -> Result<Guide> {
        let mut candidate = self.working.clone();
        let list = list_mut(&mut candidate, kind)
            .ok_or_else(|| invalid(format!("{kind:?} guide list is absent")))?;
        if index >= list.guides.len() {
            return Err(index_error(index, list.guides.len()));
        }
        let removed = list.guides.remove(index);
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(removed)
    }

    /// Remove one guide by its validated native ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
        codec::validate_value(&candidate, self.original.limits)?;
        self.working = candidate;
        Ok(true)
    }

    /// Clear one owner list while retaining the other list and its source data.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn clear(&mut self, kind: ListKind) -> Result<bool> {
        self.set_list(kind, None)
    }

    /// Validate and consume this edit into a source-checked commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn commit(self) -> Result<Prepared> {
        if !self.has_changes() {
            let patch = Change::new(self.original.clone(), self.original.clone());
            return Ok(Prepared {
                snapshot: self.original,
                patch,
                changed: false,
            });
        }
        codec::validate_value(&self.working, self.original.limits)?;
        let updated = Arc::new(codec::rewrite_source(
            self.original.source_xml(),
            &self.working,
            self.original.limits,
        )?);
        let guides = Guides::from_xml(updated.as_slice(), self.original.limits)?;
        if guides != self.working {
            return Err(invalid(
                "extended-guide serialization changed the typed value",
            ));
        }
        let snapshot = Snapshot::from_wire(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            updated,
            self.original.limits,
        )?;
        let patch = Change::new(self.original, snapshot.clone());
        Ok(Prepared {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// A successful guide edit and its reversible package patch.
#[derive(Clone, Debug)]
pub(crate) struct Prepared {
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    snapshot: Snapshot,
    patch: Change,
    #[cfg_attr(
        not(test),
        expect(dead_code, reason = "read by the module's source-level tests")
    )]
    changed: bool,
}

impl Prepared {
    #[cfg(test)]
    fn has_changes(&self) -> bool {
        self.changed
    }

    /// The outcome of this edit or commit.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.has_changes())
    }

    /// Borrow the projected post-edit snapshot.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }

    /// Consume the commit into its patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn into_patch(self) -> Change {
        self.patch
    }
}

/// A reversible source-checked replacement of one presentation XML source.
#[derive(Clone, Debug)]
pub(crate) struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    /// Source context required before publication.
    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Source context produced by publication.
    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch is an exact no-op.
    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply this patch atomically to an OPC package.
    ///
    /// A change to a signed package is refused with
    /// [`OpcError::SignedSourceRequiresExplicitPolicy`](litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
    /// (ADR 0006) before anything is mutated: remove the signature explicitly
    /// first, with [`Package::unsign`](crate::Package::unsign) or, on a raw
    /// package, [`OpcPackage::unsign`](litchi_opc::OpcPackage::unsign). An exact
    /// no-op keeps the signature.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(test)]
    pub(crate) fn apply(&self, target: &mut OpcPackage) -> Result<Snapshot> {
        super::package::apply_patch(target, self)
    }

    /// Apply this patch atomically to an owned presentation XML buffer.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[cfg(test)]
    pub(crate) fn apply_xml(&self, target: &mut Vec<u8>, limits: Limits) -> Result<Snapshot> {
        if target.as_slice() != self.before.source_xml() {
            return Err(invalid("extended-guide source is stale"));
        }
        if self.is_noop() {
            let mut snapshot = self.before.clone();
            snapshot.limits = limits;
            return Ok(snapshot);
        }
        let parsed = Snapshot::from_wire(
            self.after.presentation_part_name.clone(),
            self.after.presentation_content_type.clone(),
            Arc::clone(self.after.source_arc()),
            limits,
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
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x100_0000_01b3);
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
