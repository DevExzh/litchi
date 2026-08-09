//! Failure-atomic, source-bound edits for workbook Custom XML Maps.

use litchi_core::sheet::Result;
use litchi_opc::OpcPackage;

use super::codec;
use super::invalid;
use super::model::{XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema};
use super::package;
use super::patch::{Commit, Patch};
use super::snapshot::Snapshot;

/// A typed transaction over the workbook's single Custom XML Maps owner.
pub struct Transaction<'a> {
    target: &'a mut OpcPackage,
    before: Snapshot,
    draft: Option<XmlMapInfo>,
    conformance: XmlMapConformance,
}

impl<'a> Transaction<'a> {
    /// Start a transaction after validating and capturing the workbook graph.
    pub fn new(target: &'a mut OpcPackage) -> Result<Self> {
        let before = Snapshot::load(target)?;
        Ok(Self {
            target,
            draft: before.info().cloned(),
            conformance: before.conformance(),
            before,
        })
    }

    /// Start a transaction and select the conformance for newly published XML.
    /// Existing owners are retained in their source conformance by [`Self::new`]
    /// unless this explicit constructor is used.
    pub fn new_with_conformance(
        target: &'a mut OpcPackage,
        conformance: XmlMapConformance,
    ) -> Result<Self> {
        let mut transaction = Self::new(target)?;
        transaction.conformance = conformance;
        Ok(transaction)
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the currently staged Custom XML Maps value.
    #[must_use]
    pub fn info(&self) -> Option<&XmlMapInfo> {
        self.draft.as_ref()
    }

    /// Contextual alias for [`Self::info`].
    #[must_use]
    pub fn xml_maps(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// Contextual catalog alias for [`Self::info`].
    #[must_use]
    pub fn catalog(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// Explicit alias for the typed `MapInfo` value.
    #[must_use]
    pub fn map_info(&self) -> Option<&XmlMapInfo> {
        self.info()
    }

    /// The conformance that will be used at publication.
    #[must_use]
    pub fn conformance(&self) -> XmlMapConformance {
        self.conformance
    }

    /// Replace the complete owner, or remove it with `None`.
    pub fn replace(&mut self, value: Option<XmlMapInfo>) -> Result<bool> {
        validate_draft(value.as_ref(), self.conformance)?;
        if self.draft == value {
            return Ok(false);
        }
        self.draft = value;
        Ok(true)
    }

    /// Replace or create the complete owner.
    pub fn set(&mut self, value: XmlMapInfo) -> Result<bool> {
        self.replace(Some(value))
    }

    /// Edit the complete typed owner through a cloned, validated draft.
    pub fn edit(&mut self, edit: impl FnOnce(&mut XmlMapInfo) -> Result<()>) -> Result<bool> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent custom XML maps owner"))?;
        edit(&mut draft)?;
        validate_draft(Some(&draft), self.conformance)?;
        if self.draft.as_ref() == Some(&draft) {
            return Ok(false);
        }
        self.draft = Some(draft);
        Ok(true)
    }

    /// Select the conformance for the next publication.
    pub fn set_conformance(&mut self, conformance: XmlMapConformance) -> Result<bool> {
        if self.conformance == conformance {
            return Ok(false);
        }
        if let Some(value) = self.draft.as_ref() {
            validate_draft(Some(value), conformance)?;
        }
        self.conformance = conformance;
        Ok(self.draft.is_some())
    }

    /// Insert one schema and return its staged index.
    pub fn insert_schema(&mut self, schema: XmlMapSchema) -> Result<usize> {
        let mut draft = self.draft.clone().ok_or_else(|| {
            invalid("cannot insert a schema into an absent custom XML maps owner")
        })?;
        draft.schemas.push(schema);
        validate_draft(Some(&draft), self.conformance)?;
        let index = draft.schemas.len() - 1;
        self.draft = Some(draft);
        Ok(index)
    }

    /// Edit one schema by its stable schema ID.
    pub fn edit_schema(
        &mut self,
        id: &str,
        edit: impl FnOnce(&mut XmlMapSchema) -> Result<()>,
    ) -> Result<bool> {
        self.edit(|value| {
            let schema = value
                .schemas
                .iter_mut()
                .find(|schema| schema.id == id)
                .ok_or_else(|| invalid(format!("Schema ID '{id}' was not found")))?;
            edit(schema)
        })
    }

    /// Remove one schema by ID, retaining the transaction on validation failure.
    pub fn remove_schema(&mut self, id: &str) -> Result<Option<XmlMapSchema>> {
        let mut draft = self.draft.clone().ok_or_else(|| {
            invalid("cannot remove a schema from an absent custom XML maps owner")
        })?;
        let Some(index) = draft.schemas.iter().position(|schema| schema.id == id) else {
            return Ok(None);
        };
        let removed = draft.schemas.remove(index);
        validate_draft(Some(&draft), self.conformance)?;
        self.draft = Some(draft);
        Ok(Some(removed))
    }

    /// Insert one map and return its staged index.
    pub fn insert_map(&mut self, map: XmlMap) -> Result<usize> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot insert a map into an absent custom XML maps owner"))?;
        draft.maps.push(map);
        validate_draft(Some(&draft), self.conformance)?;
        let index = draft.maps.len() - 1;
        self.draft = Some(draft);
        Ok(index)
    }

    /// Edit one map by its stable map ID.
    pub fn edit_map(
        &mut self,
        id: u32,
        edit: impl FnOnce(&mut XmlMap) -> Result<()>,
    ) -> Result<bool> {
        self.edit(|value| {
            let map = value
                .maps
                .iter_mut()
                .find(|map| map.id == id)
                .ok_or_else(|| invalid(format!("Map ID {id} was not found")))?;
            edit(map)
        })
    }

    /// Remove one map by ID, retaining the transaction on validation failure.
    pub fn remove_map(&mut self, id: u32) -> Result<Option<XmlMap>> {
        let mut draft = self
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot remove a map from an absent custom XML maps owner"))?;
        let Some(index) = draft.maps.iter().position(|map| map.id == id) else {
            return Ok(None);
        };
        let removed = draft.maps.remove(index);
        validate_draft(Some(&draft), self.conformance)?;
        self.draft = Some(draft);
        Ok(Some(removed))
    }

    /// Remove the complete owner and return its staged semantic value.
    pub fn remove(&mut self) -> Result<Option<XmlMapInfo>> {
        Ok(self.draft.take())
    }

    /// Whether staged semantics or publication conformance differ from source.
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.before.info() != self.draft.as_ref()
            || (self.draft.is_some() && self.before.conformance() != self.conformance)
    }

    /// Validate and atomically publish the staged owner.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let patch = Patch::new(self.before.clone(), self.before.clone());
            return Ok(Commit::new(self.before, patch, false));
        }

        let current = Snapshot::load(self.target)?;
        if !current.same_source(&self.before) {
            return Err(invalid("custom XML maps transaction source is stale"));
        }

        let mut candidate = self.target.clone();
        match (self.before.info(), self.draft.as_ref()) {
            (Some(before), Some(after)) => {
                let source = self
                    .before
                    .source_xml()
                    .ok_or_else(|| invalid("custom XML maps source XML is absent"))?;
                let xml = codec::patch_source(
                    source,
                    before,
                    after,
                    self.before.conformance().is_strict(),
                    self.conformance.is_strict(),
                )?;
                package::store_xml_in_package(&mut candidate, &xml, self.conformance)?;
            },
            (None, Some(after)) => {
                package::store_in_package(&mut candidate, after, self.conformance)?;
            },
            (Some(_), None) => {
                package::remove_from_package(&mut candidate)?;
            },
            (None, None) => {},
        }

        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.info() != self.draft.as_ref()
            || (self.draft.is_some() && snapshot.conformance() != self.conformance)
        {
            return Err(invalid(
                "custom XML maps publication changed the staged model",
            ));
        }
        let patch = Patch::new(self.before, snapshot.clone());
        *self.target = candidate;
        Ok(Commit::new(snapshot, patch, true))
    }
}

fn validate_draft(value: Option<&XmlMapInfo>, conformance: XmlMapConformance) -> Result<()> {
    if let Some(value) = value {
        value.to_xml(conformance.is_strict())?;
    }
    Ok(())
}
