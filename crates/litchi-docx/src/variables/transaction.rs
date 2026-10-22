//! Source-checked snapshots, edits, commits, and patches for document variables.

use std::sync::Arc;

use crate::{Error, Result};
use litchi_core::SourceVersion;

use super::codec;
use super::model::Variables;
use super::raw::Limits;

/// An immutable settings XML snapshot with its typed document-variable view.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<Vec<u8>>,
    variables: Variables,
    source_version: Option<SourceVersion>,
    limits: Limits,
}

impl Snapshot {
    /// Parse and retain a bounded complete Word settings XML snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: Vec<u8>, limits: Limits) -> Result<Self> {
        Self::from_shared_xml(Arc::new(xml), limits)
    }

    pub(crate) fn from_shared_xml(xml: Arc<Vec<u8>>, limits: Limits) -> Result<Self> {
        let variables = Variables::from_xml(xml.as_slice(), &limits)?;
        Ok(Self {
            xml,
            variables,
            source_version: None,
            limits,
        })
    }

    pub(crate) fn from_source_xml(
        xml: Arc<Vec<u8>>,
        source_version: SourceVersion,
        limits: Limits,
    ) -> Result<Self> {
        let mut snapshot = Self::from_shared_xml(xml, limits)?;
        snapshot.source_version = Some(source_version);
        Ok(snapshot)
    }

    /// Borrow the exact settings XML retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Borrow the typed document-variable collection in source order.
    #[inline]
    #[must_use]
    pub const fn variables(&self) -> &Variables {
        &self.variables
    }

    /// Start an isolated source-checked edit.
    #[must_use]
    pub(crate) fn edit(&self) -> Edit {
        Edit {
            base: self.clone(),
            next: self.variables.clone(),
        }
    }

    /// Start an isolated edit using only fallible semantic allocation.
    ///
    /// # Errors
    ///
    /// Returns an allocation error if the bounded variable collection cannot
    /// be cloned.
    pub(crate) fn try_edit(&self) -> Result<Edit> {
        Ok(Edit {
            base: self.clone(),
            next: self.variables.try_clone()?,
        })
    }

    fn same_source(&self, other: &Self) -> bool {
        self.source_version == other.source_version
            && self.xml.as_slice() == other.xml.as_slice()
            && self.variables == other.variables
    }

    pub(crate) fn shared_xml(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }
}

/// A document-variable edit that has not yet been published.
#[derive(Debug, Clone)]
pub(crate) struct Edit {
    base: Snapshot,
    next: Variables,
}

impl Edit {
    /// Borrow the projected document-variable collection.
    #[inline]
    #[must_use]
    pub(crate) const fn variables(&self) -> &Variables {
        &self.next
    }

    /// Insert or replace one document variable atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn set(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<&mut Self> {
        self.next
            .insert_with_limits(name, value, &self.base.limits)?;
        Ok(self)
    }

    /// Remove one variable and return its previous value.
    pub(crate) fn remove(&mut self, name: &str) -> Option<String> {
        self.next.remove(name)
    }

    /// Remove every staged variable.
    pub(crate) fn clear(&mut self) -> &mut Self {
        self.next.clear();
        self
    }

    /// Replace the complete collection after validating all Word limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn replace(&mut self, value: Variables) -> Result<&mut Self> {
        value.validate_with_limits(&self.base.limits)?;
        self.next = value;
        Ok(self)
    }

    /// Apply a clone-staged collection mutation atomically.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn edit_variables(
        &mut self,
        edit: impl FnOnce(&mut Variables) -> Result<()>,
    ) -> Result<&mut Self> {
        let mut candidate = self.next.try_clone()?;
        edit(&mut candidate)?;
        candidate.validate_with_limits(&self.base.limits)?;
        self.next = candidate;
        Ok(self)
    }

    /// Validate and publish this edit without mutating its source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub(crate) fn commit(self) -> Result<Commit> {
        self.next.validate_with_limits(&self.base.limits)?;
        if self.next == self.base.variables {
            let patch = Patch {
                before: self.base.clone(),
                after: self.base.clone(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
                changed: false,
            });
        }

        let xml = codec::rewrite(
            self.base.xml_bytes(),
            &self.base.variables,
            &self.next,
            &self.base.limits,
        )?;
        let snapshot = Snapshot::from_shared_xml_with_lineage(
            xml,
            self.base.source_version,
            self.base.limits,
        )?;
        if snapshot.variables != self.next {
            return Err(invalid(
                "document-variable publication changed the staged collection",
            ));
        }
        let patch = Patch {
            before: self.base,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

impl Snapshot {
    fn from_shared_xml_with_lineage(
        xml: Vec<u8>,
        source_version: Option<SourceVersion>,
        limits: Limits,
    ) -> Result<Self> {
        let mut snapshot = Self::from_shared_xml(Arc::new(xml), limits)?;
        snapshot.source_version = source_version;
        Ok(snapshot)
    }
}

/// A successful document-variable publication.
#[derive(Debug, Clone)]
#[allow(dead_code)]
pub(crate) struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    /// Whether the private staging engine changed the source.
    #[allow(dead_code)]
    pub(crate) const fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the private staging engine's published read view.
    #[allow(dead_code)]
    pub(crate) const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the private staging engine's reversible patch.
    #[allow(dead_code)]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub(crate) fn into_patch(self) -> Patch {
        self.patch
    }

    /// Move both private staging products out of the engine result.
    #[allow(dead_code)]
    pub(crate) fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// A reversible patch with exact source and typed semantic preconditions.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Borrow the expected source collection.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Variables {
        self.before.variables()
    }

    /// Borrow the collection produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Variables {
        self.after.variables()
    }

    /// Borrow the exact source snapshot used by this patch.
    #[inline]
    #[must_use]
    pub const fn before_snapshot(&self) -> &Snapshot {
        &self.before
    }

    /// Borrow the exact snapshot produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after_snapshot(&self) -> &Snapshot {
        &self.after
    }

    /// Whether this patch leaves both source semantics and bytes unchanged.
    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before.same_source(&self.after)
    }

    /// Return the exact inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Apply to a snapshot with the patch's exact before bytes and
    /// variables, from any source (record 0355a: a package replay reads the
    /// snapshot it applies to).
    pub(crate) fn apply_content(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_slice() != self.before.xml.as_slice()
            || source.variables != self.before.variables
        {
            return Err(Error::conflict(
                "document_variables",
                "the patch base does not match the supplied source",
            ));
        }
        Ok(self.after.clone())
    }

    /// Apply only to a snapshot matching both source preconditions.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if !source.same_source(&self.before) {
            return Err(Error::conflict(
                "document_variables",
                "the patch base does not match the supplied source",
            ));
        }
        Ok(self.after.clone())
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
