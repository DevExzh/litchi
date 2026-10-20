//! Move-only source-bound slide-ID Designer-tag transactions.

use std::sync::Arc;

#[cfg(test)]
use litchi_opc::OpcPackage;

use super::codec;
use super::model::Snapshot;
use super::{Limits, Tags};
use crate::{Error, Result};

/// Stable fingerprint of the selected slide-ID host and its OPC binding.
pub type Revision = u64;

/// An isolated edit that owns its source snapshot.
#[derive(Debug)]
pub(crate) struct Draft {
    original: Snapshot,
    desired: Option<Tags>,
}

impl Draft {
    pub(crate) fn new(original: Snapshot) -> Self {
        let desired = original.occurrences.first().cloned();
        Self { original, desired }
    }

    /// Replace or create the tag list after bounded validation.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn set(&mut self, value: Tags) -> Result<bool> {
        validate_tags(&value, self.original.limits)?;
        if self.desired.as_ref() == Some(&value) {
            return Ok(false);
        }
        self.desired = Some(value);
        Ok(true)
    }

    /// Remove the singular tag-list extension.
    #[inline]
    pub(crate) fn remove(&mut self) -> bool {
        self.desired.take().is_some()
    }

    fn is_changed(&self) -> bool {
        self.original.occurrences.first() != self.desired.as_ref()
    }

    /// Consume the edit into a candidate-reparsed reversible commit.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub(crate) fn commit(self) -> Result<Prepared> {
        if !self.is_changed() {
            let before = self.original;
            let after = before.duplicate();
            return Ok(Prepared {
                patch: Change { before, after },
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
        Ok(Prepared {
            patch: Change {
                before: self.original,
                after,
            },
        })
    }
}

/// A successful candidate-reparsed edit and its reversible patch.
#[derive(Debug)]
pub(crate) struct Prepared {
    patch: Change,
}

impl Prepared {
    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub(crate) fn patch(&self) -> &Change {
        &self.patch
    }

    /// Consume the commit into its reversible patch.
    #[inline]
    #[must_use]
    #[cfg(test)]
    pub(crate) fn into_patch(self) -> Change {
        self.patch
    }
}

/// A reversible replacement bound to one exact stable slide-ID host.
#[derive(Debug)]
pub(crate) struct Change {
    pub(crate) before: Snapshot,
    pub(crate) after: Snapshot,
}

impl Change {
    fn is_changed(&self) -> bool {
        !self.before.same_selected_source(&self.after)
    }

    /// Whether this patch leaves its source unchanged.
    #[must_use]
    pub(crate) fn is_noop(&self) -> bool {
        !self.is_changed()
    }

    /// Build the exact inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(&self) -> Self {
        Self {
            before: self.after.duplicate(),
            after: self.before.duplicate(),
        }
    }

    /// Apply this patch atomically after re-resolving the stable slide ID.
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
    #[inline]
    #[cfg(test)]
    pub(crate) fn apply(&self, package: &mut OpcPackage) -> Result<Snapshot> {
        super::apply_patch(package, self)
    }
}

fn validate_tags(value: &Tags, limits: Limits) -> Result<()> {
    // Reusing the strict writer provides the model's private aggregate
    // consistency validation without duplicating its invariants here.
    crate::shape::designer::write_tags(value, limits).map(|_| ())
}
