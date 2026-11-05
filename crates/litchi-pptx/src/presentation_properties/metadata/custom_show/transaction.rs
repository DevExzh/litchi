//! Detached, source-checked custom-show transactions.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::Limits;
use super::model::{List, Show};
use super::{package, wire};
use crate::{Error, Result};

/// Stable fingerprint of the exact `PresentationML` source and relationship topology.
pub type Revision = u64;

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) presentation_part_name: String,
    pub(crate) presentation_content_type: String,
    pub(crate) source_xml: Arc<Vec<u8>>,
    pub(crate) list: List,
    pub(crate) layout: wire::Layout,
    pub(crate) relationships: Vec<wire::RelationshipState>,
    pub(crate) revision: Revision,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Load a validated source-bound custom-show snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn load(package: &OpcPackage, limits: Limits) -> Result<Self> {
        package::load_snapshot(package, limits)
    }

    pub(crate) fn from_located(
        presentation_part_name: String,
        presentation_content_type: String,
        source_xml: Arc<Vec<u8>>,
        located: wire::Located,
        limits: Limits,
    ) -> Result<Self> {
        if source_xml.len() > limits.max_input_bytes() {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::InputBytes),
                observed: source_xml.len(),
                maximum: limits.max_input_bytes(),
            });
        }
        let revision = fingerprint(
            source_xml.as_slice(),
            &located.relationships,
            &presentation_part_name,
            &presentation_content_type,
        );
        Ok(Self {
            presentation_part_name,
            presentation_content_type,
            source_xml,
            list: located.list,
            layout: located.layout,
            relationships: located.relationships,
            revision,
            limits,
        })
    }

    /// Borrow the typed custom-show list in source order.
    #[inline]
    #[must_use]
    pub fn list(&self) -> &List {
        &self.list
    }

    /// Return the owning `PresentationML` part name.
    #[inline]
    #[must_use]
    pub fn presentation_part_name(&self) -> &str {
        &self.presentation_part_name
    }

    /// Return the owning `PresentationML` content type.
    #[inline]
    #[must_use]
    pub fn presentation_content_type(&self) -> &str {
        &self.presentation_content_type
    }

    /// Borrow the exact `PresentationML` bytes captured by this snapshot.
    #[inline]
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source_xml.as_slice()
    }

    /// Return the revision used for optimistic stale-source checks.
    #[inline]
    #[must_use]
    pub const fn revision(&self) -> Revision {
        self.revision
    }

    /// Start an isolated edit over the typed custom-show list.
    #[inline]
    #[must_use]
    pub(crate) fn edit(&self) -> Draft {
        Draft {
            original: self.clone(),
            working: self.list.clone(),
        }
    }

    pub(crate) fn source_arc(&self) -> &Arc<Vec<u8>> {
        &self.source_xml
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.presentation_part_name == other.presentation_part_name
            && self.presentation_content_type == other.presentation_content_type
            && self.source_xml.as_slice() == other.source_xml.as_slice()
            && self.relationships == other.relationships
            && self.revision == other.revision
    }
}

/// A failure-atomic edit staged against one source snapshot.
#[derive(Clone, Debug)]
pub struct Draft {
    original: Snapshot,
    working: List,
}

impl Draft {
    /// Borrow the currently staged typed list.
    #[inline]
    #[must_use]
    pub fn list(&self) -> &List {
        &self.working
    }

    fn has_changes(&self) -> bool {
        self.original.list.shows != self.working.shows
    }

    /// The outcome of this edit or commit.
    #[must_use]
    pub fn diagnostics(&self) -> crate::diagnostics::Diagnostics {
        crate::diagnostics::Diagnostics::new(self.has_changes())
    }

    /// Replace the complete typed list after validating all slide references.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace(&mut self, value: List) -> Result<bool> {
        validate(&self.original, &value)?;
        if self.working.shows == value.shows {
            return Ok(false);
        }
        self.working = value;
        Ok(true)
    }

    /// Apply a checked mutation to a cloned list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn edit(&mut self, edit: impl FnOnce(&mut List) -> Result<()>) -> Result<()> {
        let mut candidate = self.working.clone();
        edit(&mut candidate)?;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Add a complete custom show.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add(&mut self, show: Show) -> Result<()> {
        self.insert(self.working.shows.len(), show)
    }

    /// Create a show using the list's next available ID and return that ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn create(&mut self, name: impl Into<String>, slide_ids: Vec<u32>) -> Result<u32> {
        let mut candidate = self.working.clone();
        let id = candidate.create(name, slide_ids).id;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(id)
    }

    /// Insert a show at a source-order position.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn insert(&mut self, index: usize, show: Show) -> Result<()> {
        if index > self.working.shows.len() {
            return Err(index_error(index, self.working.shows.len()));
        }
        let mut candidate = self.working.clone();
        candidate.shows.insert(index, show);
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Remove a show by stable ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove(&mut self, id: u32) -> Result<Show> {
        let mut candidate = self.working.clone();
        let removed = candidate
            .remove_by_id(id)
            .ok_or_else(|| invalid(format!("custom show {id} was not found")))?;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(removed)
    }

    /// Remove a show by display name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_by_name(&mut self, name: &str) -> Result<Show> {
        let mut candidate = self.working.clone();
        let removed = candidate
            .remove_by_name(name)
            .ok_or_else(|| invalid(format!("custom show '{name}' was not found")))?;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(removed)
    }

    /// Replace one show while retaining its stable ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn replace_show(&mut self, id: u32, replacement: Show) -> Result<bool> {
        let mut candidate = self.working.clone();
        let before = candidate
            .get_by_id(id)
            .ok_or_else(|| invalid(format!("custom show {id} was not found")))?
            .clone();
        candidate.replace_by_id(id, replacement)?;
        validate(&self.original, &candidate)?;
        let changed = before
            != *candidate
                .get_by_id(id)
                .ok_or_else(|| invalid(format!("custom show {id} disappeared")))?;
        self.working = candidate;
        Ok(changed)
    }

    /// Set a show's display name.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_name(&mut self, id: u32, name: impl Into<String>) -> Result<bool> {
        let mut candidate = self.working.clone();
        let show = candidate
            .get_by_id_mut(id)
            .ok_or_else(|| invalid(format!("custom show {id} was not found")))?;
        let name = name.into();
        if show.name == name {
            return Ok(false);
        }
        show.name = name;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Replace one show's slide membership in source order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set_slides(&mut self, id: u32, slide_ids: Vec<u32>) -> Result<bool> {
        let mut candidate = self.working.clone();
        let show = candidate
            .get_by_id_mut(id)
            .ok_or_else(|| invalid(format!("custom show {id} was not found")))?;
        if show.slide_ids == slide_ids {
            return Ok(false);
        }
        show.slide_ids = slide_ids;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Add one slide to a show.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn add_slide(&mut self, show_id: u32, slide_id: u32) -> Result<()> {
        let mut candidate = self.working.clone();
        let show = candidate
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        show.slide_ids.push(slide_id);
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Remove one slide from a show.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn remove_slide(&mut self, show_id: u32, slide_id: u32) -> Result<bool> {
        let mut candidate = self.working.clone();
        let show = candidate
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        let Some(index) = show.slide_ids.iter().position(|value| *value == slide_id) else {
            return Ok(false);
        };
        show.slide_ids.remove(index);
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(true)
    }

    /// Reorder custom shows by a complete ID permutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder(&mut self, ordered_ids: &[u32]) -> Result<()> {
        let mut candidate = self.working.clone();
        candidate.reorder(ordered_ids)?;
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Reorder one show's slides by a complete membership permutation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn reorder_slides(&mut self, show_id: u32, ordered_ids: &[u32]) -> Result<()> {
        let mut candidate = self.working.clone();
        let show = candidate
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        let expected = show
            .slide_ids
            .iter()
            .copied()
            .collect::<std::collections::HashSet<_>>();
        let actual = ordered_ids
            .iter()
            .copied()
            .collect::<std::collections::HashSet<_>>();
        if expected != actual || show.slide_ids.len() != ordered_ids.len() {
            return Err(invalid("custom-show slide reorder is not a permutation"));
        }
        show.slide_ids = ordered_ids.to_vec();
        validate(&self.original, &candidate)?;
        self.working = candidate;
        Ok(())
    }

    /// Consume the edit into a source-checked commit.
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
        validate(&self.original, &self.working)?;
        let source = wire::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.layout,
            &self.original.list,
            &self.working,
            self.original.limits,
        )?;
        let (list, layout) =
            wire::decode_rewritten(&source, &self.original.layout, self.original.limits)?;
        if list.shows != self.working.shows {
            return Err(invalid(
                "custom-show serialization changed the typed semantic list",
            ));
        }
        let snapshot = Snapshot::from_located(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            Arc::new(source),
            wire::Located {
                list,
                layout,
                relationships: self.original.relationships.clone(),
            },
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

/// A successful custom-show edit and its reversible patch.
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

    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }

    #[must_use]
    #[cfg(test)]
    pub(crate) fn into_patch(self) -> Change {
        self.patch
    }
}

/// A reversible source-checked replacement of the owning presentation XML.
#[derive(Clone, Debug)]
pub(crate) struct Change {
    before: Snapshot,
    after: Snapshot,
}

impl Change {
    fn new(before: Snapshot, after: Snapshot) -> Self {
        Self { before, after }
    }

    #[inline]
    #[must_use]
    pub(crate) fn before(&self) -> &Snapshot {
        &self.before
    }

    #[inline]
    #[must_use]
    pub(crate) fn after(&self) -> &Snapshot {
        &self.after
    }

    #[inline]
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        self.before.same_source(&self.after)
    }

    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

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
        package::apply_patch(target, self)
    }
}

fn validate(original: &Snapshot, value: &List) -> Result<()> {
    wire::validate_list(
        value,
        original.layout.slide_id_to_relationship.keys(),
        original.limits,
    )
}

fn fingerprint(
    source: &[u8],
    relationships: &[wire::RelationshipState],
    part_name: &str,
    content_type: &str,
) -> Revision {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for value in [source, part_name.as_bytes(), content_type.as_bytes()] {
        for byte in value {
            hash ^= u64::from(*byte);
            hash = hash.wrapping_mul(0x100_0000_01b3);
        }
    }
    for relationship in relationships {
        for value in [
            relationship.id.as_bytes(),
            relationship.relationship_type.as_bytes(),
            relationship.target.as_bytes(),
            &[u8::from(relationship.external)],
        ] {
            for byte in value {
                hash ^= u64::from(*byte);
                hash = hash.wrapping_mul(0x100_0000_01b3);
            }
        }
    }
    hash
}

fn index_error(index: usize, len: usize) -> Error {
    invalid(format!(
        "custom-show index {index} is outside a list of length {len}"
    ))
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
