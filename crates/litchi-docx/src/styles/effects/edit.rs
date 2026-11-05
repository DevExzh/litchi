//! The stylesWithEffects family (record 0355): the optional Word
//! `stylesWithEffects` part of one owner (the main or the glossary
//! document).
//!
//! [`Package::styles_with_effects`](crate::Package::styles_with_effects)
//! reads an owner's [`Snapshot`];
//! [`Package::edit_styles_with_effects`](crate::Package::edit_styles_with_effects)
//! stages an owner-bound [`Edit`] whose commit publishes a new package with
//! the owner's source-checked [`Patch`], and
//! [`Package::apply_styles_with_effects`](crate::Package::apply_styles_with_effects)
//! replays that patch onto another package holding the same before state.

use super::{Draft, Owner, Patch, Resource, Snapshot, apply_patch_staged, load};
use crate::Package;

/// What a stylesWithEffects commit changed.
pub use litchi_core::edit::Diagnostics;

/// The error of the stylesWithEffects family: the crate's error.
pub type Error = crate::Error;

/// A published stylesWithEffects change: the new package, its patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

pub use super::limit::LimitKind;

/// A staged stylesWithEffects change of one owner over one package snapshot.
/// The source is never changed: [`Self::commit`] publishes a new package,
/// and dropping the edit discards it.
#[derive(Clone)]
pub struct Edit {
    source: Package,
    owner: Owner,
    next: Option<Option<Resource>>,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("owner", &self.owner)
            .finish_non_exhaustive()
    }
}

impl Edit {
    /// The owner this edit changes.
    #[must_use]
    pub const fn owner(&self) -> Owner {
        self.owner
    }

    /// Stage the owner's resource: `Some` publishes it, `None` removes the
    /// part.
    pub fn replace_resource(&mut self, resource: Option<Resource>) -> &mut Self {
        self.next = Some(resource);
        self
    }

    /// Stage `resource` as the owner's part.
    pub fn put(&mut self, resource: Resource) -> &mut Self {
        self.replace_resource(Some(resource))
    }

    /// Stage the removal of the owner's part.
    pub fn remove(&mut self) -> &mut Self {
        self.replace_resource(None)
    }

    /// Discard the staged change, keeping the edit.
    pub fn reset(&mut self) {
        self.next = None;
    }

    /// Validate the staged resource and publish it as a new package with
    /// its source-checked patch. An unchanged resource publishes the source
    /// package unchanged.
    ///
    /// # Errors
    ///
    /// Returns an error for a resource whose conformance differs from the
    /// package, a missing glossary owner, an unpublishable graph, and the
    /// signed-package refusal for a change to a signed package.
    pub fn commit(self) -> crate::Result<Commit> {
        let current = load(self.source.opc_package(), self.owner)?;
        let mut draft: Draft = current.edit();
        if let Some(next) = self.next {
            draft.replace_resource(next)?;
        }
        let patch = draft.commit()?.into_patch();
        if patch.is_noop() {
            return Ok(Commit::new(self.source, patch, Diagnostics::UNCHANGED));
        }
        let package = publish(&self.source, self.owner, &patch, current)?;
        Ok(Commit::new(package, patch, Diagnostics::CHANGED))
    }
}

fn publish(
    source: &Package,
    owner: Owner,
    patch: &Patch,
    current: Snapshot,
) -> crate::Result<Package> {
    let mut package = source.clone();
    package.edit_semantic_opc("styles_with_effects.commit", move |opc| {
        apply_patch_staged(opc, owner, patch, current).map(|_| ())
    })?;
    Ok(package)
}

impl Package {
    /// The stylesWithEffects snapshot of `owner`; an absent part is an empty
    /// snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid stylesWithEffects graph.
    pub fn styles_with_effects(&self, owner: Owner) -> crate::Result<Snapshot> {
        load(self.opc_package(), owner)
    }

    /// Stage a stylesWithEffects change of `owner` over this package. O(1).
    #[must_use]
    pub fn edit_styles_with_effects(&self, owner: Owner) -> Edit {
        Edit {
            source: self.clone(),
            owner,
            next: None,
        }
    }

    /// Replay a stylesWithEffects patch onto its owner in this package, which
    /// must hold the patch's before state.
    ///
    /// # Errors
    ///
    /// Returns an error for a package whose owner, resource or graph does not
    /// match the patch's precondition, and the signed-package refusal for a
    /// changed patch on a signed package.
    pub fn apply_styles_with_effects(&self, patch: &Patch) -> crate::Result<Commit> {
        let current = load(self.opc_package(), patch.owner())?;
        let projected = patch.apply_to_snapshot(&current)?;
        if current.same_state(&projected) {
            return Ok(Commit::new(
                self.clone(),
                patch.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        let package = publish(self, patch.owner(), patch, current)?;
        Ok(Commit::new(package, patch.clone(), Diagnostics::CHANGED))
    }
}
