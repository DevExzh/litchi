//! Package facade for the optional Word stylesWithEffects resources.

use super::super::model::{Package, Result};
use crate::styles::effects::{self, Commit, Owner, Patch, Resource, Snapshot};

impl Package {
    /// Load the owner-wide stylesWithEffects snapshot.
    ///
    /// A missing owner resource is returned as a snapshot whose resource is
    /// None. A missing glossary document therefore remains an absence on read.
    pub fn styles_with_effects(&self, owner: Owner) -> Result<Snapshot> {
        effects::load(&self.opc, owner)
    }

    /// Add or replace one owner resource.
    ///
    /// This returns false for an exact source/conformance no-op. Glossary
    /// publication requires an already established glossary document and does
    /// not create one implicitly.
    pub fn put_styles_with_effects(&mut self, owner: Owner, resource: Resource) -> Result<bool> {
        let current = effects::load(&self.opc, owner)?;
        if current.resource().is_some_and(|value| value == &resource) {
            return Ok(false);
        }
        self.edit_semantic_opc("put_styles_with_effects", |opc| {
            effects::put(opc, owner, resource)
        })
    }

    /// Remove one owner resource.
    ///
    /// Missing resources are an exact no-op. A glossary resource can only be
    /// removed while its glossary document owner remains valid.
    pub fn remove_styles_with_effects(&mut self, owner: Owner) -> Result<bool> {
        let current = effects::load(&self.opc, owner)?;
        if current.resource().is_none() {
            return Ok(false);
        }
        self.edit_semantic_opc("remove_styles_with_effects", |opc| {
            effects::remove(opc, owner)
        })
    }

    /// Apply a source-checked stylesWithEffects patch atomically.
    pub fn apply_styles_with_effects_patch(
        &mut self,
        owner: Owner,
        patch: &Patch,
    ) -> Result<Snapshot> {
        let current = effects::load(&self.opc, owner)?;
        let projected = patch.apply_to_snapshot(&current)?;
        if current.same_state(&projected) {
            return Ok(current);
        }
        self.edit_semantic_opc("apply_styles_with_effects_patch", move |opc| {
            effects::apply_patch_staged(opc, owner, patch, current)
        })
    }

    /// Apply a committed stylesWithEffects edit atomically.
    pub fn apply_styles_with_effects_commit(
        &mut self,
        owner: Owner,
        commit: Commit,
    ) -> Result<Snapshot> {
        self.apply_styles_with_effects_patch(owner, commit.patch())
    }
}
