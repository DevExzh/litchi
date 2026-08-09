//! Move-only source-bound slide-ID Designer-tag transactions.

use std::sync::Arc;

use litchi_opc::OpcPackage;

use super::codec;
use super::model::Snapshot;
use super::{Limits, Tags};
use crate::{Error, Result};

/// Stable fingerprint of the selected slide-ID host and its OPC binding.
pub type Revision = u64;

/// An isolated edit that owns its source snapshot.
#[derive(Debug)]
pub struct Edit {
    original: Snapshot,
    desired: Option<Tags>,
}

impl Edit {
    pub(crate) fn new(original: Snapshot) -> Self {
        let desired = original.occurrences.first().cloned();
        Self { original, desired }
    }

    /// Borrow the optional staged tag list.
    #[inline]
    #[must_use]
    pub fn tags(&self) -> Option<&Tags> {
        self.desired.as_ref()
    }

    /// Replace or create the tag list after bounded validation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn set(&mut self, value: Tags) -> Result<bool> {
        validate_tags(&value, self.original.limits)?;
        if self.desired.as_ref() == Some(&value) {
            return Ok(false);
        }
        self.desired = Some(value);
        Ok(true)
    }

    /// Remove the singular tag-list extension.
    #[inline]
    pub fn remove(&mut self) -> bool {
        self.desired.take().is_some()
    }

    /// Apply a checked mutation to a cloned present tag list.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn update(&mut self, edit: impl FnOnce(&mut Tags) -> Result<()>) -> Result<()> {
        let mut candidate = self
            .desired
            .clone()
            .ok_or_else(|| Error::Invalid("Designer tag list is absent".into()))?;
        edit(&mut candidate)?;
        validate_tags(&candidate, self.original.limits)?;
        self.desired = Some(candidate);
        Ok(())
    }

    /// Return whether the staged semantic presence/value differs from source.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        self.original.occurrences.first() != self.desired.as_ref()
    }

    /// Consume the edit into a candidate-reparsed reversible commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn commit(self) -> Result<Commit> {
        if !self.is_changed() {
            let before = self.original;
            let after = before.duplicate();
            return Ok(Commit {
                patch: Patch { before, after },
                changed: false,
            });
        }
        if let Some(value) = &self.desired {
            validate_tags(value, self.original.limits)?;
        }
        let source = codec::rewrite(
            self.original.source_xml.as_slice(),
            &self.original.layout,
            self.desired.as_ref(),
            self.original.limits,
        )?;
        let located = codec::locate(&source, self.original.slide_id, self.original.limits)?;
        if located.tags.as_slice() != self.desired.as_slice() {
            return Err(Error::Invalid(
                "Designer-tag serialization changed the staged semantic value".into(),
            ));
        }
        if located.layout.relationship_id != self.original.binding.relationship_id {
            return Err(Error::Invalid(
                "Designer-tag serialization changed the slide relationship binding".into(),
            ));
        }
        let after = super::package::snapshot_from_located(
            self.original.presentation_part_name.clone(),
            self.original.presentation_content_type.clone(),
            Arc::new(source),
            self.original.slide_id,
            self.original.binding.clone(),
            located,
            self.original.limits,
        )?;
        Ok(Commit {
            patch: Patch {
                before: self.original,
                after,
            },
            changed: true,
        })
    }
}

/// A successful candidate-reparsed edit and its reversible patch.
#[derive(Debug)]
pub struct Commit {
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Return whether the commit changes the selected host.
    #[inline]
    #[must_use]
    pub const fn is_changed(&self) -> bool {
        self.changed
    }

    /// Borrow the candidate-reparsed result snapshot.
    #[inline]
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.patch.after
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its reversible patch.
    #[inline]
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }

    /// Publish this commit atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn apply(self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::apply_commit(package, self)
    }
}

/// A reversible replacement bound to one exact stable slide-ID host.
#[derive(Debug)]
pub struct Patch {
    pub(crate) before: Snapshot,
    pub(crate) after: Snapshot,
}

impl Patch {
    /// Borrow the required selected-host state.
    #[inline]
    #[must_use]
    pub fn before(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the candidate selected-host state.
    #[inline]
    #[must_use]
    pub fn after(&self) -> &Snapshot {
        &self.after
    }

    /// Return whether this patch changes the selected host.
    #[inline]
    #[must_use]
    pub fn is_changed(&self) -> bool {
        !self.before.same_selected_source(&self.after)
    }

    /// Return the required selected-host revision.
    #[inline]
    #[must_use]
    pub const fn expected_revision(&self) -> Revision {
        self.before.revision
    }

    /// Build the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.duplicate(),
            after: self.before.duplicate(),
        }
    }

    /// Apply this patch atomically after re-resolving the stable slide ID.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[inline]
    pub fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::apply_patch(package, self)
    }
}

fn validate_tags(value: &Tags, limits: Limits) -> Result<()> {
    // Reusing the strict writer provides the model's private aggregate
    // consistency validation without duplicating its invariants here.
    crate::shape::designer::write_tags(value, limits).map(|_| ())
}
