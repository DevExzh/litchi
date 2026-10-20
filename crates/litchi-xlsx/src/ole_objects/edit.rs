//! Failure-atomic, source-bound edits for worksheet OLE metadata.

use litchi_opc::{OpcPackage, PackURI};

use super::model::{OleObject, OleObjectAnchor, OleObjects};
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::Snapshot;
use super::validation;
use crate::error::{Result, invalid};

/// A staged, source-bound edit of one worksheet's OLE metadata of one
/// [`crate::Package`].
///
/// The edit changes metadata and anchors in place, and [`Self::set`] adds a
/// collection to a worksheet that has none. Shape IDs, relationships,
/// preview identity, and opaque payload bytes of an existing collection are
/// deliberately immutable so the commit preserves source markup and never
/// executes or decodes embedded content. Dropping the edit discards it;
/// [`Self::commit`] publishes a new package.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    worksheet: PackURI,
    before: Snapshot,
    draft: Option<OleObjects>,
    /// The candidate graph holding a collection added by [`Self::set`].
    created: Option<OpcPackage>,
}

impl crate::Package {
    /// Start an edit of one worksheet's OLE metadata.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, a `NotFound` refusal for a selector that matches no sheet,
    /// [`crate::Error::NotWorksheet`] for a chart or dialog sheet, and the
    /// validation errors of the worksheet's OLE graph.
    pub fn edit_ole_objects<'a>(
        &self,
        sheet: impl Into<crate::workbook::Selector<'a>>,
    ) -> Result<Edit> {
        let target = crate::record::Source::new(self, "edit_ole_objects")?;
        let worksheet = self.worksheet_part(sheet, "edit_ole_objects")?;
        Edit::start(target, &worksheet)
    }

    /// Apply an exact OLE-object patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_ole_objects(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_ole_objects")
    }
}

crate::record::raw_test_route!(edit_ole_objects, worksheet: &PackURI => crate::record::sheet_of);

impl Edit {
    fn start(target: crate::record::Source, worksheet: &PackURI) -> Result<Self> {
        let before = Snapshot::load(&target, worksheet)?;
        let draft = before.objects().cloned();
        Ok(Self {
            target,
            worksheet: worksheet.clone(),
            before,
            draft,
            created: None,
        })
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged typed graph.
    #[must_use]
    pub fn objects(&self) -> Option<&OleObjects> {
        self.draft.as_ref()
    }

    /// Edit one object by its stable `DrawingML` shape ID.
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

    /// Add a new OLE collection and its inert relationships to a worksheet
    /// that has none, in the worksheet's namespace conformance.
    ///
    /// # Errors
    ///
    /// Refuses a worksheet that already has OLE objects, and returns the
    /// validation errors of the collection.
    pub fn set(&mut self, value: OleObjects) -> Result<bool> {
        if self.draft.is_some() {
            return Err(invalid("worksheet already contains OLE objects"));
        }
        let mut candidate = OpcPackage::clone(&self.target);
        super::package::store_ole_objects(
            &mut candidate,
            &self.worksheet,
            &value,
            self.before.conformance(),
        )?;
        let stored = Snapshot::load(&candidate, &self.worksheet)?;
        self.draft = stored.objects().cloned();
        self.created = Some(candidate);
        Ok(true)
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.draft = self.before.objects().cloned();
        self.created = None;
    }

    fn changed(&self) -> bool {
        self.created.is_some() || self.before.objects() != self.draft.as_ref()
    }

    /// Validate and publish the staged source-preserving edit as a new
    /// package. An unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when the source changed,
    /// [`crate::Error::Signed`] for a changed edit of a signed workbook, and
    /// the publication's validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    pub(super) fn publish(self) -> Result<(Commit, Snapshot)> {
        if !self.changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                self.before,
            ));
        }
        let after = self
            .draft
            .as_ref()
            .ok_or_else(|| invalid("cannot remove a worksheet OLE collection in this edit"))?;
        let created = self.created.is_some();
        let mut candidate = match self.created {
            Some(candidate) => candidate,
            None => {
                let candidate = OpcPackage::clone(&self.target);
                let current = Snapshot::load(&candidate, &self.worksheet)?;
                if !current.same_source(&self.before) {
                    return Err(crate::error::Error::PatchConflict {
                        part: self.worksheet.to_string(),
                    });
                }
                candidate
            },
        };
        let base = Snapshot::load(&candidate, &self.worksheet)?;
        if base.objects() != Some(after) {
            let objects = base
                .objects()
                .ok_or_else(|| invalid("cannot publish an absent worksheet OLE collection"))?;
            let xml = super::codec::patch_ole_objects_source(
                base.source_xml(),
                objects,
                after,
                base.conformance(),
            )?;
            candidate.get_part_mut(&self.worksheet)?.set_blob(xml);
            validation::graph(&candidate, &self.worksheet)?;
        }
        let snapshot = Snapshot::load(&candidate, &self.worksheet)?;
        if snapshot.objects() != Some(after) {
            return Err(invalid("OLE source publication changed the staged model"));
        }
        let patch = if created {
            Patch::with_parts(
                self.before,
                snapshot.clone(),
                crate::part_edit::PartPatch::capture_all(&self.target, &candidate)?,
            )
        } else {
            Patch::new(self.before, snapshot.clone())
        };
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}
