//! Failure-atomic, source-bound edits for workbook Custom XML Maps.

use crate::error::Result;
use litchi_opc::OpcPackage;

use super::codec;
use super::invalid;
use super::model::{XmlMap, XmlMapConformance, XmlMapInfo, XmlMapSchema};
use super::package;
use super::patch::{Commit, Diagnostics, Patch};
use super::snapshot::Snapshot;

/// The configuration of a Custom XML Maps edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    conformance: Option<XmlMapConformance>,
}

impl EditOptions {
    /// Publish in `conformance` instead of the existing owner's source
    /// conformance.
    #[must_use]
    pub const fn conformance(mut self, conformance: XmlMapConformance) -> Self {
        self.conformance = Some(conformance);
        self
    }
}

/// A staged, source-bound edit of the workbook's single Custom XML Maps
/// owner of one [`crate::Package`]. Dropping it discards it;
/// [`Self::commit`] publishes a new package. Starting one retains only the
/// source package and options; the graph and semantic draft are materialized
/// by the first getter, mutation or commit.
#[derive(Debug, Clone)]
pub struct Edit {
    target: crate::record::Source,
    selected_conformance: Option<XmlMapConformance>,
    loaded: Option<Loaded>,
}

#[derive(Debug, Clone)]
struct Loaded {
    before: Snapshot,
    draft: Option<XmlMapInfo>,
    conformance: XmlMapConformance,
}

impl Edit {
    fn start(target: crate::record::Source, options: EditOptions) -> Result<Self> {
        Ok(Self {
            target,
            selected_conformance: options.conformance,
            loaded: None,
        })
    }

    fn ensure_loaded(&mut self) -> Result<&mut Loaded> {
        if self.loaded.is_none() {
            let before = Snapshot::load(&self.target)?;
            let conformance = self
                .selected_conformance
                .unwrap_or_else(|| before.conformance());
            self.loaded = Some(Loaded {
                draft: before.info().cloned(),
                conformance,
                before,
            });
        }
        Ok(self.loaded.as_mut().expect("custom XML maps graph loaded"))
    }

    /// Immutable source snapshot used for conflict checks and inverse patches.
    pub fn before(&mut self) -> Result<&Snapshot> {
        Ok(&self.ensure_loaded()?.before)
    }

    /// Borrow the currently staged Custom XML Maps value.
    pub fn info(&mut self) -> Result<Option<&XmlMapInfo>> {
        Ok(self.ensure_loaded()?.draft.as_ref())
    }

    /// The conformance that will be used at publication.
    pub fn conformance(&mut self) -> Result<XmlMapConformance> {
        Ok(self.ensure_loaded()?.conformance)
    }

    /// Replace the complete owner, or remove it with `None`.
    pub fn replace(&mut self, value: Option<XmlMapInfo>) -> Result<bool> {
        let conformance = self.ensure_loaded()?.conformance;
        validate_draft(value.as_ref(), conformance)?;
        if self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            == value
        {
            return Ok(false);
        }
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = value;
        Ok(true)
    }

    /// Replace or create the complete owner.
    pub fn set(&mut self, value: XmlMapInfo) -> Result<bool> {
        self.replace(Some(value))
    }

    /// Edit the complete typed owner through a cloned, validated draft.
    pub fn edit(&mut self, edit: impl FnOnce(&mut XmlMapInfo) -> Result<()>) -> Result<bool> {
        let mut draft = self
            .ensure_loaded()?
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot edit an absent custom XML maps owner"))?;
        edit(&mut draft)?;
        let conformance = self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .conformance;
        validate_draft(Some(&draft), conformance)?;
        if self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            .as_ref()
            == Some(&draft)
        {
            return Ok(false);
        }
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = Some(draft);
        Ok(true)
    }

    /// Select the conformance for the next publication.
    pub fn set_conformance(&mut self, conformance: XmlMapConformance) -> Result<bool> {
        let loaded = self.ensure_loaded()?;
        if loaded.conformance == conformance {
            return Ok(false);
        }
        if let Some(value) = loaded.draft.as_ref() {
            validate_draft(Some(value), conformance)?;
        }
        loaded.conformance = conformance;
        Ok(loaded.draft.is_some())
    }

    /// Insert one schema and return its staged index.
    pub fn insert_schema(&mut self, schema: XmlMapSchema) -> Result<usize> {
        let conformance = self.ensure_loaded()?.conformance;
        let mut draft = self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            .clone()
            .ok_or_else(|| {
                invalid("cannot insert a schema into an absent custom XML maps owner")
            })?;
        draft.schemas.push(schema);
        validate_draft(Some(&draft), conformance)?;
        let index = draft.schemas.len() - 1;
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = Some(draft);
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

    /// Remove one schema by ID, retaining the edit on validation failure.
    pub fn remove_schema(&mut self, id: &str) -> Result<Option<XmlMapSchema>> {
        let conformance = self.ensure_loaded()?.conformance;
        let mut draft = self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            .clone()
            .ok_or_else(|| {
                invalid("cannot remove a schema from an absent custom XML maps owner")
            })?;
        let Some(index) = draft.schemas.iter().position(|schema| schema.id == id) else {
            return Ok(None);
        };
        let removed = draft.schemas.remove(index);
        validate_draft(Some(&draft), conformance)?;
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = Some(draft);
        Ok(Some(removed))
    }

    /// Insert one map and return its staged index.
    pub fn insert_map(&mut self, map: XmlMap) -> Result<usize> {
        let conformance = self.ensure_loaded()?.conformance;
        let mut draft = self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot insert a map into an absent custom XML maps owner"))?;
        draft.maps.push(map);
        validate_draft(Some(&draft), conformance)?;
        let index = draft.maps.len() - 1;
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = Some(draft);
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

    /// Remove one map by ID, retaining the edit on validation failure.
    pub fn remove_map(&mut self, id: u32) -> Result<Option<XmlMap>> {
        let conformance = self.ensure_loaded()?.conformance;
        let mut draft = self
            .loaded
            .as_ref()
            .expect("custom XML maps graph loaded")
            .draft
            .clone()
            .ok_or_else(|| invalid("cannot remove a map from an absent custom XML maps owner"))?;
        let Some(index) = draft.maps.iter().position(|map| map.id == id) else {
            return Ok(None);
        };
        let removed = draft.maps.remove(index);
        validate_draft(Some(&draft), conformance)?;
        self.loaded
            .as_mut()
            .expect("custom XML maps graph loaded")
            .draft = Some(draft);
        Ok(Some(removed))
    }

    /// Remove the complete owner and return its staged semantic value.
    pub fn remove(&mut self) -> Result<Option<XmlMapInfo>> {
        Ok(self.ensure_loaded()?.draft.take())
    }

    /// Discard every staged change, keeping the publication conformance.
    pub fn reset(&mut self) -> Result<()> {
        let loaded = self.ensure_loaded()?;
        loaded.draft = loaded.before.info().cloned();
        Ok(())
    }

    fn changed(loaded: &Loaded) -> bool {
        loaded.before.info() != loaded.draft.as_ref()
            || (loaded.draft.is_some() && loaded.before.conformance() != loaded.conformance)
    }

    /// Validate and publish the staged owner as a new package.
    ///
    /// An edit without a change returns the source package, so a signed
    /// package keeps its signature byte for byte. A changed commit over a
    /// signed package is refused with [`crate::Error::Signed`]: remove the
    /// signature explicitly first ([`crate::Package::unsign`]).
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::Signed`], a stale-source error and the
    /// publication's validation errors.
    pub fn commit(self) -> Result<Commit> {
        self.publish().map(|(commit, _)| commit)
    }

    /// [`Self::commit`], keeping the published snapshot.
    pub(super) fn publish(mut self) -> Result<(Commit, Snapshot)> {
        let unchanged = {
            let loaded = self.ensure_loaded()?;
            !Self::changed(loaded)
        };
        if unchanged {
            let loaded = self.loaded.take().expect("custom XML maps graph loaded");
            let patch = Patch::new(loaded.before.clone(), loaded.before.clone());
            return Ok((
                Commit::new(self.target.into_package(), patch, Diagnostics::new(false)),
                loaded.before,
            ));
        }
        let loaded = self.loaded.take().expect("custom XML maps graph loaded");
        let Loaded {
            before,
            draft,
            conformance,
        } = loaded;
        crate::signature_policy::ensure_unsigned(&self.target)?;

        let current = Snapshot::load(&self.target)?;
        if !current.same_source(&before) {
            return Err(invalid("custom XML maps edit source is stale"));
        }

        let mut candidate = OpcPackage::clone(&self.target);
        match (before.info(), draft.as_ref()) {
            (Some(before_info), Some(after)) => {
                let source = before
                    .source_xml()
                    .ok_or_else(|| invalid("custom XML maps source XML is absent"))?;
                let xml = codec::patch_source(
                    source,
                    before_info,
                    after,
                    before.conformance().is_strict(),
                    conformance.is_strict(),
                )?;
                package::store_xml_in_package(&mut candidate, &xml, conformance)?;
            },
            (None, Some(after)) => {
                package::store_in_package(&mut candidate, after, conformance)?;
            },
            (Some(_), None) => {
                package::remove_from_package(&mut candidate)?;
            },
            (None, None) => {},
        }

        let snapshot = Snapshot::load(&candidate)?;
        if snapshot.info() != draft.as_ref()
            || (draft.is_some() && snapshot.conformance() != conformance)
        {
            return Err(invalid(
                "custom XML maps publication changed the staged model",
            ));
        }
        let patch = Patch::new(before, snapshot.clone());
        let package = self.target.publish(candidate)?;
        Ok((
            Commit::new(package, patch, Diagnostics::new(true)),
            snapshot,
        ))
    }
}

impl crate::Package {
    /// Start an edit of the workbook's Custom XML Maps.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused. Custom XML Maps graph, XML and source-limit errors are
    /// deferred to the first getter, mutation or commit.
    pub fn edit_xml_maps(&self) -> Result<Edit> {
        self.edit_xml_maps_with_options(EditOptions::default())
    }

    /// Start a Custom XML Maps edit with explicit options.
    ///
    /// # Errors
    ///
    /// See [`Self::edit_xml_maps`].
    pub fn edit_xml_maps_with_options(&self, options: EditOptions) -> Result<Edit> {
        Edit::start(crate::record::Source::new(self, "edit_xml_maps")?, options)
    }

    /// Apply an exact Custom XML Maps patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`crate::Error::PatchConflict`] when this package is not the
    /// patch's before state and [`crate::Error::Signed`] for a changed patch
    /// of a signed workbook.
    pub fn apply_xml_maps(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_xml_maps")
    }
}

crate::record::raw_test_route!(edit_xml_maps);

fn validate_draft(value: Option<&XmlMapInfo>, conformance: XmlMapConformance) -> Result<()> {
    if let Some(value) = value {
        value.to_xml(conformance.is_strict())?;
    }
    Ok(())
}
