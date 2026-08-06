//! Failure-atomic, source-bound edits for worksheet OLE metadata.

use litchi_opc::{OpcPackage, PackURI};

use super::model::{OleObject, OleObjectAnchor, OleObjects};
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;
use super::validation;
use crate::error::{Result, invalid};

/// A typed transaction over one worksheet's existing OLE metadata.
///
/// The transaction edits metadata and anchors in place. Shape IDs,
/// relationships, preview identity, and opaque payload bytes are deliberately
/// immutable here so the commit can preserve source markup and never execute
/// or decode embedded content. New/removal topology remains available through
/// the lower-level package storage API.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    worksheet: PackURI,
    before: Snapshot,
    draft: Option<OleObjects>,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating and capturing the worksheet graph.
    pub fn new(target: &'a mut OpcPackage, worksheet: &PackURI) -> Result<Self> {
        let before = Snapshot::load(target, worksheet)?;
        let draft = before.objects().cloned();
        Ok(Self {
            target,
            worksheet: worksheet.clone(),
            before,
            draft,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged typed graph.
    pub fn objects(&self) -> Option<&OleObjects> {
        self.draft.as_ref()
    }

    /// Edit one object by its stable DrawingML shape ID.
    ///
    /// The closure is applied to a private clone and validated before it
    /// becomes staged state. An error leaves the transaction unchanged.
    pub fn edit_object(
        &mut self,
        shape_id: u32,
        edit: impl FnOnce(&mut OleObject) -> Result<()>,
    ) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent worksheet OLE collection"))?;
        let object = draft
            .objects
            .iter_mut()
            .find(|object| object.shape_id == shape_id)
            .ok_or_else(|| invalid(format!("OLE shapeId {shape_id} is absent")))?;
        edit(object)?;
        if object.shape_id != shape_id {
            return Err(invalid(
                "OLE shapeId is immutable inside a metadata transaction",
            ));
        }
        validation::objects(&draft)?;
        if self.draft.as_ref() == Some(&draft) {
            return Ok(false);
        }
        self.draft = Some(draft);
        Ok(true)
    }

    /// Replace one object's typed anchor while preserving its payload and IDs.
    pub fn set_anchor(&mut self, shape_id: u32, anchor: OleObjectAnchor) -> Result<bool> {
        self.edit_object(shape_id, |object| {
            let properties = object
                .properties
                .as_mut()
                .ok_or_else(|| invalid("OLE object has no objectPr anchor"))?;
            properties.anchor = anchor;
            Ok(())
        })
    }

    /// Whether staged typed metadata differs from the captured source model.
    pub fn is_changed(&self) -> bool {
        self.before.objects() != self.draft.as_ref()
    }

    /// Validate and publish the staged source-preserving edit atomically.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }
        let before = self
            .before
            .objects()
            .ok_or_else(|| invalid("cannot publish an absent worksheet OLE collection"))?;
        let after = self.draft.as_ref().ok_or_else(|| {
            invalid("cannot remove a worksheet OLE collection in this transaction")
        })?;

        let mut candidate = self.target.clone();
        let current = Snapshot::load(&candidate, &self.worksheet)?;
        if !current.same_source(&self.before) {
            return Err(crate::error::Error::PatchConflict {
                part: self.worksheet.to_string(),
            });
        }
        let xml = super::codec::patch_ole_objects_source(
            current.source_xml(),
            before,
            after,
            current.conformance(),
        )?;
        candidate.get_part_mut(&self.worksheet)?.set_blob(xml);
        validation::graph(&candidate, &self.worksheet)?;
        let snapshot = Snapshot::load(&candidate, &self.worksheet)?;
        if snapshot.objects() != Some(after) {
            return Err(invalid("OLE source publication changed the staged model"));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}
